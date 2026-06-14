<#
.SYNOPSIS
    Open every user's calendar to the whole organisation (run AFTER historical events are masked).

.DESCRIPTION
    Sets the "Default" (= everyone in the organisation) permission on each user's Calendar folder.
    The Default principal is what every internal user sees, so this is how you make calendars
    visible org-wide. Source: Set-MailboxFolderPermission (Exchange Online PowerShell).

    Choose -AccessRight:
      AvailabilityOnly  Free/busy only (no subjects at all, past or future). Most private.
      LimitedDetails    Free/busy + subject + location (NO body).
      Reviewer          Full details (read-only) = OWA "Can view all details".

    Private items stay masked at every one of these levels, because this script does NOT grant the
    ViewPrivateItems sharing flag. So once historical events are marked Private (see
    Set-HistoricalCalendarPrivate.ps1), and the owner marks future sensitive meetings Private,
    those remain hidden while everything else is visible.
    Sources: KB 4021947; Set-MailboxFolderPermission -SharingPermissionFlags (ViewPrivateItems).

.PARAMETER AccessRight
    AvailabilityOnly | LimitedDetails | Reviewer. Default: LimitedDetails.

.PARAMETER Mailbox
    Optional specific UPNs; omit to process ALL user mailboxes.

.NOTES
    Requires: PowerShell 7+, ExchangeOnlineManagement module.
    Run after Set-HistoricalCalendarPrivate.ps1. Test with -WhatIf first.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet("AvailabilityOnly","LimitedDetails","Reviewer")]
    [string]$AccessRight = "LimitedDetails",

    [string[]]$Mailbox
)

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw "ExchangeOnlineManagement not found. Install: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -ShowBanner:$false

if ($Mailbox) {
    $boxes = $Mailbox | ForEach-Object { Get-Mailbox -Identity $_ -ErrorAction Stop }
} else {
    $boxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
}

$count = 0
foreach ($mb in $boxes) {
    $calId = "$($mb.PrimarySmtpAddress):\Calendar"   # English folder name. See note below for localised tenants.
    if ($PSCmdlet.ShouldProcess($calId, "Set Default = $AccessRight (ViewPrivateItems NOT granted)")) {
        try {
            # Default entry always exists; Set- modifies it. Explicitly clear private-item visibility.
            Set-MailboxFolderPermission -Identity $calId -User Default `
                -AccessRights $AccessRight -SharingPermissionFlags None -ErrorAction Stop
            Write-Host "OK   $($mb.PrimarySmtpAddress) -> $AccessRight" -ForegroundColor Green
            $count++
        }
        catch {
            Write-Warning "FAIL $($mb.PrimarySmtpAddress): $($_.Exception.Message)"
        }
    }
}
Write-Host "`nDONE. Calendars updated: $count" -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false

# NOTE on localised mailboxes: the Calendar folder may be named in the mailbox's language
# (e.g. "Kalender" in Swedish). If ":\Calendar" errors, resolve the localised folder name per
# mailbox with Get-MailboxFolderStatistics -FolderScope Calendar, then build the identity from
# that DisplayName.
