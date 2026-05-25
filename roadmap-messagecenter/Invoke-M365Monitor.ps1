#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft 365 Message Center & Roadmap Monitor

.DESCRIPTION
    Collects items from Message Center (Graph API) and the M365 Roadmap (public feed),
    compares them against a local JSON database, emails new/changed items, then saves
    the database atomically. A .bak copy is kept so the database survives script failures.

    DB is only written AFTER a successful email send, so a transient send failure
    automatically retries on the next run.

.PARAMETER TenantId        Azure AD Tenant ID
.PARAMETER ClientId        App registration Client ID
.PARAMETER ClientSecret    App registration Client Secret (use a secure vault in production)
.PARAMETER SenderMailbox   Mailbox UPN to send from (app needs Mail.Send permission)
.PARAMETER Recipients      One or more recipient email addresses
.PARAMETER NoEmail         Collect and update DB without sending email
.PARAMETER ForceFullReport Send email even when no changes are detected

.EXAMPLE
    .\Invoke-M365Monitor.ps1 `
        -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientId     "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -ClientSecret "your-secret" `
        -SenderMailbox "reports@contoso.com" `
        -Recipients    "admin@contoso.com","team@contoso.com"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]   $TenantId,
    [Parameter(Mandatory)] [string]   $ClientId,
    [Parameter(Mandatory)] [string]   $ClientSecret,
    [Parameter(Mandatory)] [string]   $SenderMailbox,
    [Parameter(Mandatory)] [string[]] $Recipients,
    [switch] $NoEmail,
    [switch] $ForceFullReport
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:DbPath = Join-Path $PSScriptRoot "m365-monitor.json"
$GRAPH         = "https://graph.microsoft.com/v1.0"

# ── Logging ────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $color = switch ($Level) { "ERROR" {"Red"} "WARN" {"Yellow"} "SUCCESS" {"Green"} default {"Cyan"} }
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')][$Level] $Msg" -ForegroundColor $color
}

# ── Database ───────────────────────────────────────────────────────────────

function Read-Db {
    $toHt = { param($obj) $h = @{}; $obj?.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }; $h }

    foreach ($f in @($script:DbPath, "$($script:DbPath).bak")) {
        if (-not (Test-Path $f)) { continue }
        try {
            $obj = Get-Content $f -Raw -Encoding UTF8 | ConvertFrom-Json
            Write-Log "DB loaded from '$(Split-Path $f -Leaf)' (MC: $((&$toHt $obj.messageCenter).Count), Roadmap: $((&$toHt $obj.roadmap).Count))"
            return @{
                lastRun       = $obj.lastRun
                messageCenter = & $toHt $obj.messageCenter
                roadmap       = & $toHt $obj.roadmap
            }
        } catch { Write-Log "DB '$f' unreadable: $_" "WARN" }
    }

    Write-Log "No usable database found — starting fresh" "WARN"
    return @{ lastRun = $null; messageCenter = @{}; roadmap = @{} }
}

function Save-Db {
    param($Db)
    $tmp = "$($script:DbPath).tmp"
    $bak = "$($script:DbPath).bak"

    # Write to temp first — if this fails, the current DB is untouched
    $Db | ConvertTo-Json -Depth 10 | Set-Content $tmp -Encoding UTF8 -Force

    # Rotate backup
    if (Test-Path $script:DbPath) { Copy-Item $script:DbPath $bak -Force }

    # Atomic rename (single filesystem operation on NTFS/ext4)
    Move-Item $tmp $script:DbPath -Force
    Write-Log "DB saved → $(Split-Path $script:DbPath -Leaf)"
}

# ── Graph ──────────────────────────────────────────────────────────────────

function Get-Token {
    $r = Invoke-RestMethod -Method POST `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -Body @{
            grant_type    = "client_credentials"
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = "https://graph.microsoft.com/.default"
        }
    return $r.access_token
}

function Get-GraphPaged {
    param([string]$Uri, [string]$Token)
    $items = [System.Collections.Generic.List[object]]::new()
    $url   = $Uri
    do {
        $r = Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $Token" }
        if ($r.value) { $items.AddRange([object[]]$r.value) }
        $url = $r.'@odata.nextLink'
    } while ($url)
    return $items
}

# ── Message Center ─────────────────────────────────────────────────────────

function Get-MessageCenter {
    param([string]$Token)
    # Requires app permission: ServiceMessage.Read.All
    return Get-GraphPaged `
        -Uri "$GRAPH/admin/serviceAnnouncement/messages?`$top=100&`$orderby=lastModifiedDateTime desc" `
        -Token $Token
}

# ── Roadmap ────────────────────────────────────────────────────────────────

function Get-Roadmap {
    # The M365 Roadmap has no official API — data is extracted from the public page's
    # Next.js SSR payload (__NEXT_DATA__). If parsing breaks, check the page source at:
    # https://www.microsoft.com/en-us/microsoft-365/roadmap
    try {
        $html = (Invoke-WebRequest `
            -Uri "https://www.microsoft.com/en-us/microsoft-365/roadmap" `
            -UseBasicParsing -TimeoutSec 30).Content

        if ($html -match '<script id="__NEXT_DATA__"[^>]*>([^<]+)</script>') {
            $nd = $Matches[1] | ConvertFrom-Json
            # Try known paths in the Next.js data tree
            foreach ($path in @("props.pageProps.items", "props.pageProps.roadmapItems", "props.pageProps.features")) {
                $node = $nd
                $ok   = $true
                foreach ($seg in $path.Split('.')) {
                    if ($null -eq $node -or -not $node.PSObject.Properties[$seg]) { $ok = $false; break }
                    $node = $node.$seg
                }
                if ($ok -and $node) {
                    Write-Log "Roadmap: $($node.Count) item(s) via __NEXT_DATA__.$path"
                    return $node
                }
            }
        }

        Write-Log "Roadmap: could not parse page JSON — page structure may have changed" "WARN"
        return @()
    } catch {
        Write-Log "Roadmap fetch failed: $_" "WARN"
        return @()
    }
}

# ── Diff ───────────────────────────────────────────────────────────────────

function Get-McDiff {
    param($Items, $Db)
    $new = [System.Collections.Generic.List[object]]::new()
    $upd = [System.Collections.Generic.List[object]]::new()
    foreach ($i in $Items) {
        $prev = $Db.messageCenter[$i.id]
        if     (-not $prev)                                                    { $new.Add($i) }
        elseif ($prev.lastModifiedDateTime -ne $i.lastModifiedDateTime)        { $upd.Add($i) }
    }
    return @{ New = $new; Updated = $upd }
}

function Get-RmDiff {
    param($Items, $Db)
    $new = [System.Collections.Generic.List[object]]::new()
    $upd = [System.Collections.Generic.List[object]]::new()
    foreach ($i in $Items) {
        $id   = [string]$i.id
        $prev = $Db.roadmap[$id]
        if (-not $prev) {
            $new.Add($i)
        } elseif ($prev.status -ne $i.status) {
            $i | Add-Member -NotePropertyName prevStatus -NotePropertyValue $prev.status -Force
            $upd.Add($i)
        }
    }
    return @{ New = $new; Updated = $upd }
}

# ── Email ──────────────────────────────────────────────────────────────────

function Get-TagString {
    param($Tags)
    if (-not $Tags) { return "" }
    return ($Tags | ForEach-Object { if ($_ -is [string]) { $_ } else { $_.name } }) -join ", "
}

function Build-EmailHtml {
    param($McDiff, $RmDiff)

    $ts     = (Get-Date).ToUniversalTime().ToString("dddd, MMMM d yyyy HH:mm 'UTC'")
    $mcTotal = $McDiff.New.Count + $McDiff.Updated.Count
    $rmTotal = $RmDiff.New.Count + $RmDiff.Updated.Count

    function New-Pill { param([string]$Text, [string]$Bg)
        "<span style='background:$Bg;color:#fff;border-radius:3px;padding:2px 8px;font-size:.75em;font-weight:600;white-space:nowrap'>$Text</span>"
    }
    function New-Row { param([string]$Pill, [string]$Title, [string]$Meta)
        "<tr><td style='padding:9px 4px;border-bottom:1px solid #ececec;vertical-align:top'>$Pill&nbsp; <strong>$Title</strong><br><span style='color:#777;font-size:.83em'>$Meta</span></td></tr>"
    }

    $mcRows = ""
    foreach ($i in $McDiff.New) {
        $mod  = if ($i.lastModifiedDateTime) { $i.lastModifiedDateTime.Substring(0,10) } else { "-" }
        $mcRows += New-Row (New-Pill "NEW" "#107c10") $i.title "$($i.id) &bull; $($i.messageType) &bull; severity: $($i.severity) &bull; $mod"
    }
    foreach ($i in $McDiff.Updated) {
        $mod  = if ($i.lastModifiedDateTime) { $i.lastModifiedDateTime.Substring(0,10) } else { "-" }
        $mcRows += New-Row (New-Pill "UPDATED" "#ca5010") $i.title "$($i.id) &bull; $($i.messageType) &bull; severity: $($i.severity) &bull; $mod"
    }

    $rmRows = ""
    foreach ($i in $RmDiff.New) {
        $rmRows += New-Row (New-Pill "NEW" "#0078d4") $i.title "ID $($i.id) &bull; $($i.status) &bull; $(Get-TagString $i.tags)"
    }
    foreach ($i in $RmDiff.Updated) {
        $pill   = New-Pill "$($i.prevStatus) → $($i.status)" "#8764b8"
        $rmRows += New-Row $pill $i.title "ID $($i.id) &bull; $(Get-TagString $i.tags)"
    }

    $mcSection = if ($mcRows) {
        "<h2 style='color:#005a9e;border-bottom:2px solid #0078d4;padding-bottom:5px;margin-top:28px'>Message Center ($mcTotal)</h2><table style='width:100%;border-collapse:collapse'>$mcRows</table>"
    } else { "" }

    $rmSection = if ($rmRows) {
        "<h2 style='color:#005a9e;border-bottom:2px solid #0078d4;padding-bottom:5px;margin-top:28px'>Roadmap ($rmTotal)</h2><table style='width:100%;border-collapse:collapse'>$rmRows</table>"
    } else { "" }

    $noChange = if (-not $mcRows -and -not $rmRows) {
        "<p style='color:#888;font-style:italic;margin-top:20px'>No changes detected since last run.</p>"
    } else { "" }

    return @"
<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="font-family:'Segoe UI',system-ui,sans-serif;max-width:860px;margin:0 auto;padding:28px;color:#1f1f1f">
<h1 style="margin:0 0 4px;color:#0078d4;font-size:1.6em">M365 Monitor</h1>
<p style="color:#888;margin:0 0 8px;font-size:.88em">$ts</p>
$mcSection$rmSection$noChange
<hr style="border:none;border-top:1px solid #e8e8e8;margin-top:36px">
<p style="color:#bbb;font-size:.75em;margin:8px 0 0">Invoke-M365Monitor.ps1</p>
</body></html>
"@
}

function Send-Email {
    param([string]$Token, [string]$Html, $McDiff, $RmDiff)

    $mcTotal = $McDiff.New.Count + $McDiff.Updated.Count
    $rmTotal = $RmDiff.New.Count + $RmDiff.Updated.Count
    $subject = "M365 Monitor — MC: $mcTotal | Roadmap: $rmTotal | $(Get-Date -Format 'yyyy-MM-dd')"

    $body = @{
        message = @{
            subject      = $subject
            body         = @{ contentType = "HTML"; content = $Html }
            toRecipients = @($Recipients | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
        }
        saveToSentItems = $false
    } | ConvertTo-Json -Depth 10 -Compress

    Invoke-RestMethod -Method POST `
        -Uri "$GRAPH/users/$SenderMailbox/sendMail" `
        -Headers @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json" } `
        -Body $body | Out-Null

    Write-Log "Email sent → $($Recipients -join ', ')" "SUCCESS"
}

# ── Update DB ──────────────────────────────────────────────────────────────

function Update-Db {
    param($Db, $McItems, $RmItems)
    $now = (Get-Date).ToUniversalTime().ToString("o")

    foreach ($i in $McItems) {
        $Db.messageCenter[$i.id] = @{
            title                = $i.title
            messageType          = $i.messageType
            severity             = $i.severity
            lastModifiedDateTime = $i.lastModifiedDateTime
            firstSeen            = $Db.messageCenter[$i.id]?.firstSeen ?? $now
        }
    }

    foreach ($i in $RmItems) {
        $id = [string]$i.id
        $Db.roadmap[$id] = @{
            title     = $i.title
            status    = $i.status
            tags      = Get-TagString $i.tags
            firstSeen = $Db.roadmap[$id]?.firstSeen ?? $now
        }
    }

    $Db.lastRun = $now
}

# ── Main ───────────────────────────────────────────────────────────────────

Write-Log "M365 Monitor starting — DB: $($script:DbPath)"

$db      = Read-Db
$token   = Get-Token
Write-Log "Token acquired" "SUCCESS"

$mcItems = Get-MessageCenter -Token $token
Write-Log "Message Center: $($mcItems.Count) item(s)"

$rmItems = Get-Roadmap
Write-Log "Roadmap: $($rmItems.Count) item(s)"

$mcDiff = Get-McDiff -Items $mcItems -Db $db
$rmDiff = Get-RmDiff -Items $rmItems -Db $db

$total  = $mcDiff.New.Count + $mcDiff.Updated.Count + $rmDiff.New.Count + $rmDiff.Updated.Count
Write-Log "Changes — MC: $($mcDiff.New.Count) new / $($mcDiff.Updated.Count) updated | Roadmap: $($rmDiff.New.Count) new / $($rmDiff.Updated.Count) status changes"

# Email — DB is intentionally written AFTER this block.
# If Send-Email throws, $ErrorActionPreference=Stop exits here and the DB stays
# unchanged, so the next run will retry sending the same items.
if (-not $NoEmail -and ($total -gt 0 -or $ForceFullReport)) {
    $html = Build-EmailHtml -McDiff $mcDiff -RmDiff $rmDiff
    Send-Email -Token $token -Html $html -McDiff $mcDiff -RmDiff $rmDiff
} elseif ($NoEmail) {
    Write-Log "-NoEmail set; skipping send" "WARN"
} else {
    Write-Log "No changes — no email sent"
}

Update-Db  -Db $db -McItems $mcItems -RmItems $rmItems
Save-Db    -Db $db

Write-Log "Done" "SUCCESS"
