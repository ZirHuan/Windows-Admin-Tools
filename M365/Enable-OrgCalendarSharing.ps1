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

    Private items stay masked at every one of these levels. Exchange Online does not allow
    SharingPermissionFlags on Default/Anonymous entries, but private items are always hidden
    from Default access regardless. Sources: KB 4021947.

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

$ScriptVersion = '1.2.0'
Write-Host "Enable-OrgCalendarSharing v$ScriptVersion" -ForegroundColor Cyan

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
    try {
        # Resolve the actual calendar folder name (varies by mailbox language, e.g. "Kalender" in Swedish)
        $calFolder = Get-MailboxFolderStatistics -Identity $mb.PrimarySmtpAddress `
            -FolderScope Calendar -ErrorAction Stop |
            Where-Object { $_.FolderType -eq 'Calendar' } |
            Select-Object -First 1
        if (-not $calFolder) {
            Write-Warning "SKIP $($mb.PrimarySmtpAddress): Calendar folder not found via FolderStatistics"
            continue
        }
        $calId = "$($mb.PrimarySmtpAddress):$($calFolder.FolderPath -replace '/', '\')"
    }
    catch {
        Write-Warning "SKIP $($mb.PrimarySmtpAddress): Could not resolve calendar folder - $($_.Exception.Message)"
        continue
    }

    if ($PSCmdlet.ShouldProcess($calId, "Set Default = $AccessRight (ViewPrivateItems NOT granted)")) {
        try {
            Set-MailboxFolderPermission -Identity $calId -User Default `
                -AccessRights $AccessRight -ErrorAction Stop
            Write-Host "OK   $($mb.PrimarySmtpAddress) [$($calFolder.FolderPath)] -> $AccessRight" -ForegroundColor Green
            $count++
        }
        catch {
            Write-Warning "FAIL $($mb.PrimarySmtpAddress): $($_.Exception.Message)"
        }
    }
}
Write-Host "`nDONE. Calendars updated: $count" -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
