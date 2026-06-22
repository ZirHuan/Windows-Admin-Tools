#Requires -Version 7.0
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
    Entra tenant ID. Required when using app-only auth (-ClientId / -ClientSecret).

.PARAMETER ClientId
    App registration client ID for app-only auth.

.PARAMETER ClientSecret
    Client secret value for app-only auth.

.NOTES
    Requires: PowerShell 7+, Microsoft.Graph.Authentication, Microsoft.Graph.Users modules.
    Auth: pass -TenantId/-ClientId/-ClientSecret for app-only (Calendars.ReadWrite application
    permission required). Omit for interactive delegated auth.
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

$ScriptVersion = '1.2.0'
Write-Host "Set-HistoricalCalendarPrivate v$ScriptVersion" -ForegroundColor Cyan

# --- Connect ---
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    throw "Microsoft.Graph.Authentication module not found. Install: Install-Module Microsoft.Graph -Scope CurrentUser"
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop

if ($TenantId -and $ClientId -and $ClientSecret) {
    $SecureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    $Credential   = New-Object System.Management.Automation.PSCredential($ClientId, $SecureSecret)
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $Credential -NoWelcome
} else {
    Connect-MgGraph -Scopes "Calendars.ReadWrite","User.Read.All" -NoWelcome
}

$cutoffString = $CutoffUtc.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
Write-Host "Cutoff (UTC): $cutoffString" -ForegroundColor Cyan

# --- Resolve mailbox list ---
if ($Mailbox) {
    $users = $Mailbox | ForEach-Object { Get-MgUser -UserId $_ -ErrorAction Stop }
} else {
    $users = Get-MgUser -All -Filter "accountEnabled eq true" -Property "id,userPrincipalName,mail" |
             Where-Object { $_.Mail }
}

$totalPatched = 0
foreach ($u in $users) {
    Write-Host "`n=== $($u.UserPrincipalName) ===" -ForegroundColor Yellow
    $patched = 0

    # Fetch events via raw Graph call — no Microsoft.Graph.Calendar module needed
    try {
        $filterEncoded = [uri]::EscapeDataString("end/dateTime lt '$cutoffString'")
        $uri = "https://graph.microsoft.com/v1.0/users/$($u.Id)/events?`$filter=$filterEncoded&`$select=id,subject,sensitivity,type,end&`$top=999"
        $events = [System.Collections.Generic.List[object]]::new()
        do {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject -ErrorAction Stop
            foreach ($ev in $response.value) { $events.Add($ev) }
            $uri = $response.'@odata.nextLink'
        } while ($uri)
    }
    catch {
        Write-Warning "  Could not read calendar for $($u.UserPrincipalName): $($_.Exception.Message)"
        continue
    }

    Write-Host "  Found: $($events.Count) event(s) before cutoff" -ForegroundColor Gray

    foreach ($e in $events) {
        if ($e.sensitivity -eq 'private') { continue }
        if ($SkipRecurringMasters -and $e.type -eq 'seriesMaster') {
            Write-Host "  skip (recurring master): $($e.subject)" -ForegroundColor DarkGray
            continue
        }
        if ($PSCmdlet.ShouldProcess("$($u.UserPrincipalName): $($e.subject)", "Set sensitivity=private")) {
            try {
                Invoke-MgGraphRequest -Method PATCH `
                    -Uri "https://graph.microsoft.com/v1.0/users/$($u.Id)/events/$($e.id)" `
                    -Body @{ sensitivity = "private" } `
                    -ErrorAction Stop
                $patched++
                Start-Sleep -Milliseconds $ThrottleDelayMs
            }
            catch {
                Write-Warning "  PATCH failed for '$($e.subject)': $($_.Exception.Message)"
            }
        }
    }
    Write-Host "  Masked: $patched event(s)" -ForegroundColor Green
    $totalPatched += $patched
}

Write-Host "`nDONE. Total events masked as Private: $totalPatched" -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
