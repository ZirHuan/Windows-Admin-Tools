<#
.SYNOPSIS
    Mask historical calendar events as Private before opening calendars org-wide.

.DESCRIPTION
    There is NO native way in Exchange Online to share "future events only". Calendar folder
    permissions always apply to the whole Calendar folder (past + future). To stop historical
    meetings being exposed once you open calendars, this script bulk-sets sensitivity = "private"
    on every event that ENDS before a cutoff date, across all (or selected) user mailboxes.

    Private items are masked from anyone holding normal calendar-sharing permissions: viewers see
    a busy block labelled "Private" with no subject/body/location. Source: Microsoft KB 4021947 and
    Graph event.sensitivity (normal|personal|private|confidential).

    IMPORTANT: "Private" is an honour-based Outlook/EXO flag, NOT encryption. A user granted Full
    Access, or Delegate + "view private items" (ViewPrivateItems), or an app with Calendars.Read
    application permission, can still read these items. For genuinely sensitive content use
    sensitivity labels with encryption (Microsoft Purview) instead.

.PARAMETER CutoffUtc
    Events ending before this UTC datetime are masked. e.g. (Get-Date '2026-06-14T00:00:00Z').

.PARAMETER Mailbox
    Optional specific UPNs. Omit to process ALL enabled, mailbox-enabled users.

.PARAMETER SkipRecurringMasters
    Skip recurring series masters (a long-running series may still have future occurrences you do
    NOT want masked). Off by default = recurring series are masked too.

.PARAMETER ThrottleDelayMs
    Pause between PATCH calls to ease Graph throttling.

.PARAMETER TenantId
    Entra tenant (directory) ID. Supply together with ClientId + ClientSecret to connect app-only.
    App-only is REQUIRED to write into another user's mailbox (delegated only reaches your own).

.PARAMETER ClientId
    App registration (application) ID with Graph APPLICATION permission Calendars.ReadWrite + User.Read.All.

.PARAMETER ClientSecret
    Client secret value for the app registration. App-only auth is used when all three are supplied.

.NOTES
    Requires: PowerShell 7+, Microsoft.Graph module.
    Graph scope: Calendars.ReadWrite. To touch OTHER mailboxes use APPLICATION permission (app-only)
    via -TenantId/-ClientId/-ClientSecret. Without them it falls back to delegated (own mailbox only).
    RUN AGAINST A PILOT MAILBOX FIRST with -WhatIf.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [datetime]$CutoffUtc,

    [string[]]$Mailbox,

    [switch]$SkipRecurringMasters,

    [int]$ThrottleDelayMs = 200,

    [string]$TenantId,

    [string]$ClientId,

    [string]$ClientSecret
)

$ScriptVersion = '1.1.0'
Write-Host "Set-HistoricalCalendarPrivate v$ScriptVersion" -ForegroundColor Cyan

# --- Connect ---
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Calendar)) {
    throw "Microsoft.Graph module not found. Install: Install-Module Microsoft.Graph -Scope CurrentUser"
}
Import-Module Microsoft.Graph.Calendar -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop

if ($TenantId -and $ClientId -and $ClientSecret) {
    Write-Host "Connecting app-only (application permissions)..." -ForegroundColor Cyan
    $sec  = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential($ClientId, $sec)
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $cred -NoWelcome
}
elseif (-not (Get-MgContext)) {
    Write-Host "Connecting delegated (own mailbox only)..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "Calendars.ReadWrite","User.Read.All" -NoWelcome
}

$cutoffString = $CutoffUtc.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
Write-Host "Cutoff (UTC): $cutoffString" -ForegroundColor Cyan

# --- Resolve mailbox list ---
if ($Mailbox) {
    $users = $Mailbox | ForEach-Object { Get-MgUser -UserId $_ -ErrorAction Stop }
} else {
    $users = Get-MgUser -All -Filter "accountEnabled eq true" -Property "id,userPrincipalName,mail" |
             Where-Object { $_.Mail }   # mailbox-enabled only
}

$totalPatched = 0
foreach ($u in $users) {
    Write-Host "`n=== $($u.UserPrincipalName) ===" -ForegroundColor Yellow
    $patched = 0
    try {
        # Events that END before cutoff. end/dateTime is a UTC string -> lexical compare is valid.
        $filter = "end/dateTime lt '$cutoffString'"
        $events = Get-MgUserEvent -UserId $u.Id -Filter $filter -All `
                    -Property "id,subject,sensitivity,type,recurrence,end" -ErrorAction Stop
    }
    catch {
        Write-Warning "  Could not read calendar for $($u.UserPrincipalName): $($_.Exception.Message)"
        continue
    }

    foreach ($e in $events) {
        if ($e.Sensitivity -eq 'private') { continue }                       # already masked
        if ($SkipRecurringMasters -and $e.Type -eq 'seriesMaster') {
            Write-Host "  skip (recurring master): $($e.Subject)" -ForegroundColor DarkGray
            continue
        }
        if ($PSCmdlet.ShouldProcess("$($u.UserPrincipalName): $($e.Subject)", "Set sensitivity=private")) {
            try {
                Update-MgUserEvent -UserId $u.Id -EventId $e.Id -Sensitivity "private" -ErrorAction Stop
                $patched++
                Start-Sleep -Milliseconds $ThrottleDelayMs
            }
            catch {
                Write-Warning "  PATCH failed for '$($e.Subject)': $($_.Exception.Message)"
            }
        }
    }
    Write-Host "  Masked: $patched event(s)" -ForegroundColor Green
    $totalPatched += $patched
}

Write-Host "`nDONE. Total events masked as Private: $totalPatched" -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
