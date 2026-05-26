#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft 365 Message Center & Roadmap Monitor

.DESCRIPTION
    Collects Message Center (Graph API) and Roadmap (public page) items, diffs against a
    local JSON database, emails new/changed items, and saves state atomically.

.PARAMETER TenantId        Azure AD Tenant ID (GUID)
.PARAMETER ClientId        App registration Client ID (GUID)
.PARAMETER ClientSecret    App client secret. Prefer env var M365_CLIENT_SECRET to avoid
                           the secret appearing in the process list / shell history.
.PARAMETER SenderMailbox   UPN of the mailbox to send from (app needs Mail.Send)
.PARAMETER Recipients      One or more recipient addresses
.PARAMETER NoEmail         Collect and update DB without sending
.PARAMETER ForceFullReport Send email even when no changes detected

.EXAMPLE
    $env:M365_CLIENT_SECRET = "…"
    .\Invoke-M365Monitor.ps1 -TenantId "…" -ClientId "…" `
        -SenderMailbox "reports@contoso.com" -Recipients "admin@contoso.com"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string] $TenantId,

    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string] $ClientId,

    # Prefer env var M365_CLIENT_SECRET — passing a secret on the command line exposes it
    # in the process list and shell history. The parameter is kept for non-interactive use
    # behind a vault wrapper or scheduled-task "run as" account.
    [Parameter(Mandatory = $false)]
    [string] $ClientSecret = "",

    [Parameter(Mandatory)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string] $SenderMailbox,

    [Parameter(Mandatory)]
    [ValidateCount(1, 50)]
    [string[]] $Recipients,

    [switch] $NoEmail,
    [switch] $ForceFullReport
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:DbPath  = Join-Path $PSScriptRoot "m365-monitor.json"
$GRAPH          = "https://graph.microsoft.com/v1.0"
$GRAPH_HOST     = "graph.microsoft.com"          # used to validate nextLink
$MAX_PAGE_ITEMS = 5000                            # safety cap on paged results

# ── Resolve secret ─────────────────────────────────────────────────────────
# Prefer env var so the value never appears in the process list.
if (-not $ClientSecret) {
    $ClientSecret = $env:M365_CLIENT_SECRET
}
if (-not $ClientSecret) {
    throw "Provide -ClientSecret or set the M365_CLIENT_SECRET environment variable."
}

# Validate recipient addresses
foreach ($r in $Recipients) {
    if ($r -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        throw "Invalid recipient address: '$r'"
    }
}

# ── Logging ────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $color = switch ($Level) { "ERROR" {"Red"} "WARN" {"Yellow"} "SUCCESS" {"Green"} default {"Cyan"} }
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')][$Level] $Msg" -ForegroundColor $color
}

# ── HTML encoding ──────────────────────────────────────────────────────────
# All untrusted data interpolated into HTML must pass through this function.

function ConvertTo-HtmlEncode {
    param([string]$s)
    if (-not $s) { return '' }
    $s = $s.Replace('&',  '&amp;')
    $s = $s.Replace('<',  '&lt;')
    $s = $s.Replace('>',  '&gt;')
    $s = $s.Replace('"',  '&quot;')
    $s = $s.Replace("'",  '&#39;')
    return $s
}

# ── Retry with exponential backoff ─────────────────────────────────────────

function Invoke-WithRetry {
    param([scriptblock]$Action, [int]$MaxAttempts = 3)
    $attempt = 0
    while ($true) {
        $attempt++
        try { return (& $Action) }
        catch {
            $code = 0
            try { $code = [int]$_.Exception.Response.StatusCode } catch {}

            $retryable = $code -in @(429, 500, 502, 503, 504) -or
                         ($code -eq 0 -and $_.Exception -is [System.Net.Http.HttpRequestException])

            if ($attempt -ge $MaxAttempts -or -not $retryable) { throw }

            $delay = [Math]::Pow(2, $attempt)   # 2 s, 4 s
            try {
                $ra = $_.Exception.Response.Headers['Retry-After']
                if ($ra -match '^\d+$') { $delay = [Math]::Min([int]$ra, 120) }
            } catch {}

            Write-Log "HTTP $code — retry $attempt/$MaxAttempts in ${delay}s" "WARN"
            Start-Sleep -Seconds $delay
        }
    }
}

# ── Database ───────────────────────────────────────────────────────────────

function Read-Db {
    $toHt = { param($obj) $h = @{}; $obj?.PSObject.Properties | ForEach-Object { $h[$_.Name] = $_.Value }; $h }

    foreach ($f in @($script:DbPath, "$($script:DbPath).bak")) {
        if (-not (Test-Path $f)) { continue }
        try {
            $obj = Get-Content $f -Raw -Encoding UTF8 | ConvertFrom-Json
            Write-Log "DB loaded from '$(Split-Path $f -Leaf)'"
            return @{
                lastRun       = $obj.lastRun
                messageCenter = & $toHt $obj.messageCenter
                roadmap       = & $toHt $obj.roadmap
            }
        } catch {
            Write-Log "DB '$f' unreadable — $($_.Exception.Message)" "WARN"
        }
    }

    Write-Log "No usable database found — starting fresh" "WARN"
    return @{ lastRun = $null; messageCenter = @{}; roadmap = @{} }
}

function Save-Db {
    param($Db)
    $tmp = "$($script:DbPath).tmp"
    $bak = "$($script:DbPath).bak"

    # Write to .tmp first — if this step fails, the live DB is untouched
    $Db | ConvertTo-Json -Depth 10 | Set-Content $tmp -Encoding UTF8 -Force

    # Rotate backup before overwriting live file
    if (Test-Path $script:DbPath) { Copy-Item $script:DbPath $bak -Force }

    # Atomic rename on same volume (single FS operation)
    Move-Item $tmp $script:DbPath -Force
    Write-Log "DB saved → $(Split-Path $script:DbPath -Leaf)"
}

# ── Graph auth ─────────────────────────────────────────────────────────────

function Get-Token {
    $secret = $ClientSecret          # local copy; cleared below
    try {
        $r = Invoke-WithRetry {
            Invoke-RestMethod -Method POST -TimeoutSec 30 `
                -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
                -Body @{
                    grant_type    = "client_credentials"
                    client_id     = $ClientId
                    client_secret = $secret
                    scope         = "https://graph.microsoft.com/.default"
                }
        }
        return $r.access_token
    } finally {
        # Best-effort clear — PS strings are immutable so the GC controls lifetime,
        # but removing the variable avoids accidental leaks in debug output.
        Remove-Variable -Name secret -ErrorAction SilentlyContinue
    }
}

# ── Paged Graph fetch ──────────────────────────────────────────────────────

function Get-GraphPaged {
    param([string]$Uri, [string]$Token)
    $h     = @{ Authorization = "Bearer $Token" }
    $items = [System.Collections.Generic.List[object]]::new()
    $url   = $Uri

    do {
        # Validate nextLink — must stay on graph.microsoft.com to prevent open-redirect
        $host = ([System.Uri]$url).Host
        if ($host -ne $GRAPH_HOST) {
            throw "Unexpected nextLink host '$host' — aborting pagination."
        }

        $r = Invoke-WithRetry {
            Invoke-RestMethod -Uri $url -Headers $h -TimeoutSec 60
        }

        if ($r.value) { $items.AddRange([object[]]$r.value) }

        if ($items.Count -ge $MAX_PAGE_ITEMS) {
            Write-Log "Paging cap ($MAX_PAGE_ITEMS) reached — stopping early" "WARN"
            break
        }

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
    # No official API — data extracted from the Next.js SSR payload (__NEXT_DATA__).
    # If this breaks, inspect page source at https://www.microsoft.com/en-us/microsoft-365/roadmap
    $maxResponseBytes = 10 * 1024 * 1024   # 10 MB safety cap

    try {
        $resp = Invoke-WithRetry {
            Invoke-WebRequest -Uri "https://www.microsoft.com/en-us/microsoft-365/roadmap" `
                -UseBasicParsing -TimeoutSec 30
        }

        if ($resp.RawContentLength -gt $maxResponseBytes) {
            Write-Log "Roadmap: response too large ($($resp.RawContentLength) bytes) — skipping" "WARN"
            return @()
        }

        $html = $resp.Content

        if ($html -match '<script id="__NEXT_DATA__"[^>]*>([^<]{1,5000000})</script>') {
            $nd = $Matches[1] | ConvertFrom-Json -ErrorAction Stop
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
        Write-Log "Roadmap fetch failed: $($_.Exception.Message)" "WARN"
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
        if     (-not $prev)                                               { $new.Add($i) }
        elseif ($prev.lastModifiedDateTime -ne $i.lastModifiedDateTime)   { $upd.Add($i) }
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

    $ts      = (Get-Date).ToUniversalTime().ToString("dddd, MMMM d yyyy HH:mm 'UTC'")
    $mcTotal = $McDiff.New.Count + $McDiff.Updated.Count
    $rmTotal = $RmDiff.New.Count + $RmDiff.Updated.Count

    # $Text is already HTML-encoded by the caller — kept raw inside the span
    function New-Pill { param([string]$Text, [string]$Bg)
        "<span style='background:$Bg;color:#fff;border-radius:3px;padding:2px 8px;font-size:.75em;font-weight:600;white-space:nowrap'>$Text</span>"
    }
    # $Title and $Meta must be HTML-encoded before being passed here
    function New-Row { param([string]$Pill, [string]$Title, [string]$Meta)
        "<tr><td style='padding:9px 4px;border-bottom:1px solid #ececec;vertical-align:top'>$Pill&nbsp;<strong>$Title</strong><br><span style='color:#777;font-size:.83em'>$Meta</span></td></tr>"
    }

    $mcRows = ""
    foreach ($i in $McDiff.New + $McDiff.Updated) {
        $isNew   = $McDiff.New -contains $i
        $pill    = New-Pill (if ($isNew) { "NEW" } else { "UPDATED" }) (if ($isNew) { "#107c10" } else { "#ca5010" })
        $title   = ConvertTo-HtmlEncode $i.title
        $id      = ConvertTo-HtmlEncode $i.id
        $type    = ConvertTo-HtmlEncode $i.messageType
        $sev     = ConvertTo-HtmlEncode $i.severity
        $rawMod  = if ($i.lastModifiedDateTime -and $i.lastModifiedDateTime.Length -ge 10) { $i.lastModifiedDateTime.Substring(0,10) } else { $i.lastModifiedDateTime }
        $mod     = ConvertTo-HtmlEncode $rawMod
        $mcRows += New-Row $pill $title "$id &bull; $type &bull; severity: $sev &bull; $mod"
    }

    $rmRows = ""
    foreach ($i in $RmDiff.New) {
        $pill    = New-Pill "NEW" "#0078d4"
        $title   = ConvertTo-HtmlEncode $i.title
        $id      = ConvertTo-HtmlEncode ([string]$i.id)
        $status  = ConvertTo-HtmlEncode $i.status
        $tags    = ConvertTo-HtmlEncode (Get-TagString $i.tags)
        $rmRows += New-Row $pill $title "ID $id &bull; $status &bull; $tags"
    }
    foreach ($i in $RmDiff.Updated) {
        $prev    = ConvertTo-HtmlEncode $i.prevStatus
        $cur     = ConvertTo-HtmlEncode $i.status
        $pill    = New-Pill "$prev → $cur" "#8764b8"
        $title   = ConvertTo-HtmlEncode $i.title
        $id      = ConvertTo-HtmlEncode ([string]$i.id)
        $tags    = ConvertTo-HtmlEncode (Get-TagString $i.tags)
        $rmRows += New-Row $pill $title "ID $id &bull; $tags"
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

    Invoke-WithRetry {
        Invoke-RestMethod -Method POST -TimeoutSec 60 `
            -Uri "$GRAPH/users/$SenderMailbox/sendMail" `
            -Headers @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json" } `
            -Body $body | Out-Null
    }

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

# DB is written AFTER a successful send so a transient failure retries on the next run.
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

# Clear token from session memory
Remove-Variable -Name token -ErrorAction SilentlyContinue

Write-Log "Done" "SUCCESS"
