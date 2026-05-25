# M365 Message Center & Roadmap Monitor

Collects items from the **Microsoft 365 Message Center** (Graph API) and the **M365 Roadmap** (public page), diffs them against a local JSON database, and sends an HTML email with new and changed items.

---

## Requirements

| | |
|---|---|
| PowerShell | 7.0 or later |
| App Registration | Azure AD app with the permissions below |
| Sender mailbox | Any mailbox in the tenant (shared mailbox recommended) |

### App registration permissions (Application, not Delegated)

| Permission | Purpose |
|---|---|
| `ServiceMessage.Read.All` | Read Message Center items |
| `Mail.Send` | Send email from the specified mailbox |

---

## Usage

```powershell
.\Invoke-M365Monitor.ps1 `
    -TenantId      "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientId      "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -ClientSecret  "your-secret" `
    -SenderMailbox "reports@contoso.com" `
    -Recipients    "admin@contoso.com", "team@contoso.com"
```

### Parameters

| Parameter | Required | Description |
|---|---|---|
| `-TenantId` | Yes | Azure AD Tenant ID |
| `-ClientId` | Yes | App registration Client ID |
| `-ClientSecret` | Yes | App registration Client Secret |
| `-SenderMailbox` | Yes | UPN of the mailbox to send from |
| `-Recipients` | Yes | One or more recipient addresses |
| `-NoEmail` | No | Collect and update DB, skip sending |
| `-ForceFullReport` | No | Send email even when nothing changed |

### First run

On the first run the database is empty, so **all current items are treated as new**. Use `-ForceFullReport` to get a complete snapshot, or `-NoEmail` to seed the database silently:

```powershell
# Seed the DB without spamming recipients
.\Invoke-M365Monitor.ps1 ... -NoEmail

# From the next run onwards, only deltas are emailed
.\Invoke-M365Monitor.ps1 ...
```

---

## Database & Recovery

The script stores state in `m365-monitor.json` next to the script file. Writes are atomic:

```
m365-monitor.json      ← live database
m365-monitor.json.bak  ← previous version (rotated on every write)
m365-monitor.json.tmp  ← transient write target (deleted after rename)
```

On startup, if `m365-monitor.json` is missing or corrupt, the script automatically falls back to `.bak`. If both are unreadable it starts fresh with an empty database.

The database is written **only after a successful email send**. If the send fails, the script exits with an error and the database is unchanged — the next run will retry the same items.

---

## Email behaviour

- **No changes** → no email (unless `-ForceFullReport`)
- **Changes found** → one HTML email per run with:
  - Message Center: new items (green) + updated items (orange, detected via `lastModifiedDateTime`)
  - Roadmap: new items (blue) + status changes (purple, shows `old → new`)
- `saveToSentItems` is set to `false` — emails are not kept in Sent Items

---

## Scheduling

### Windows Task Scheduler

```powershell
$action  = New-ScheduledTaskAction -Execute "pwsh.exe" `
    -Argument '-NonInteractive -File "C:\Scripts\m365-monitor\Invoke-M365Monitor.ps1" -TenantId "..." -ClientId "..." -ClientSecret "..." -SenderMailbox "reports@contoso.com" -Recipients "admin@contoso.com"'
$trigger = New-ScheduledTaskTrigger -Daily -At "08:00"
Register-ScheduledTask -TaskName "M365 Monitor" -Action $action -Trigger $trigger -RunLevel Highest
```

> Store the client secret in a Windows Credential Manager entry or Azure Key Vault instead of passing it on the command line.

---

## Roadmap data source

The M365 Roadmap has no official API. The script fetches `https://www.microsoft.com/en-us/microsoft-365/roadmap` and extracts the embedded Next.js JSON payload (`__NEXT_DATA__`). If Microsoft changes the page structure the roadmap section will log a warning and skip gracefully — Message Center continues to work normally.

To fix a broken roadmap parser: open the page in a browser, view source, search for `__NEXT_DATA__`, and update the path navigation in `Get-Roadmap` inside the script.
