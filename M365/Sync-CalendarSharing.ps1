#Requires -Version 7.0
<#
.SYNOPSIS
    Synkroniserar kalenderdelning (LimitedDetails) i Exchange Online utifrån två gruppers medlemskap.

.DESCRIPTION
    Modell: två statiska säkerhetsgrupper styr vem som ser vems kalender.
      - "Se"-gruppen    = tittare (får se andras kalendrar)
      - "Delad"-gruppen = kalendrar som exponeras

    För varje tittare (Se-medlem) sätts behörigheten LimitedDetails på varje
    Delad-medlems kalender. LimitedDetails visar titel + plats + ledig/upptagen,
    men INTE brödtext/deltagare, och RESPEKTERAR "private"-flaggan (dvs historiska
    events som maskerats med Set-HistoricalCalendarPrivate förblir dolda).

    Skriptet expanderar grupperna och sätter INDIVIDUELLA behörigheter per person
    (inte gruppbaserat) för förutsägbarhet och enkel revidering.

    RECONCILE (försiktig modell): skriptet rör endast LimitedDetails-poster som
    tillhör aktuella eller tidigare Se-medlemmar. Det RÖR ALDRIG andra
    behörighetsnivåer (t.ex. Reviewer/Editor som satts manuellt) eller poster för
    användare utanför Se-gruppens hantering. När en tittare tas bort ur Se-gruppen
    river nästa körning deras LimitedDetails-poster.

    "Private" är ett hedersflagg, INTE kryptering. Full Access, delegat med
    ViewPrivateItems, eller app med Calendars.Read (application) kan fortfarande läsa
    maskerat innehåll. För genuint känsligt: Microsoft Purview sensitivity labels.

.PARAMETER SeGroup
    DisplayName eller UPN/mail för "Se"-gruppen (tittare).

.PARAMETER DeladGroup
    DisplayName eller UPN/mail för "Delad"-gruppen (kalendrar som exponeras).

.PARAMETER PermissionLevel
    Behörighetsnivå. Standard: LimitedDetails (titel + plats). Andra: AvailabilityOnly, Reviewer.

.PARAMETER AppOnly
    Använd app-only-auth (certifikat) i stället för interaktiv inloggning.

.PARAMETER AppId
    App registration (client) ID för app-only.

.PARAMETER CertThumbprint
    Certifikatets thumbprint för app-only.

.PARAMETER Organization
    Tenant-domän för app-only, t.ex. antilopgroup.onmicrosoft.com.

.PARAMETER WhatIf
    Torrkörning - visar vad som skulle ändras utan att ändra något.

.NOTES
    Kräver: PowerShell 7+, ExchangeOnlineManagement-modulen.
    Behörigheter: kontot/appen måste kunna köra Add/Set/Remove-MailboxFolderPermission
    och Get-MailboxFolderStatistics för alla postlådor i Delad-gruppen.
    KÖR MED -WhatIf FÖRST.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$SeGroup,

    [Parameter(Mandatory = $true)]
    [string]$DeladGroup,

    [ValidateSet('LimitedDetails','AvailabilityOnly','Reviewer')]
    [string]$PermissionLevel = 'LimitedDetails',

    [switch]$AppOnly,
    [string]$AppId,
    [string]$CertThumbprint,
    [string]$Organization
)

$ScriptVersion = '1.0.1'
Write-Host "Sync-CalendarSharing v$ScriptVersion" -ForegroundColor Cyan
Write-Host "Behörighetsnivå: $PermissionLevel" -ForegroundColor Cyan

# --- Modulkontroll ---
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw "ExchangeOnlineManagement-modulen saknas. Installera: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# --- Anslut ---
try {
    if ($AppOnly) {
        if (-not ($AppId -and $CertThumbprint -and $Organization)) {
            throw "App-only kräver -AppId, -CertThumbprint och -Organization."
        }
        Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertThumbprint -Organization $Organization -ShowBanner:$false -ErrorAction Stop
    } else {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    }
}
catch {
    throw "Kunde inte ansluta till Exchange Online: $($_.Exception.Message)"
}

# --- Hjälpfunktion: expandera en grupp till en lista av primary SMTP ---
function Get-GroupMemberSmtp {
    param([string]$GroupIdentity)

    $members = @()
    # Prova mail-enabled (distribution / mail-enabled security) först,
    # fall tillbaka till Unified (M365) grupper.
    try {
        $members = Get-DistributionGroupMember -Identity $GroupIdentity -ResultSize Unlimited -ErrorAction Stop
    }
    catch {
        try {
            $members = Get-UnifiedGroupLinks -Identity $GroupIdentity -LinkType Members -ResultSize Unlimited -ErrorAction Stop
        }
        catch {
            throw "Kunde inte läsa medlemmar i gruppen '$GroupIdentity': $($_.Exception.Message)"
        }
    }

    $smtp = foreach ($m in $members) {
        if ($m.PrimarySmtpAddress) { [string]$m.PrimarySmtpAddress }
    }
    return $smtp | Where-Object { $_ } | Sort-Object -Unique
}

# --- Hjälpfunktion: hitta kalendermappens namn (svenska "Kalender" vs "Calendar") ---
function Get-CalendarFolderPath {
    param([string]$Mailbox)
    try {
        $cal = Get-MailboxFolderStatistics -Identity $Mailbox -FolderScope Calendar -ErrorAction Stop |
               Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
        if ($cal) {
            return "${Mailbox}:\$($cal.Name)"
        }
    }
    catch {
        Write-Warning "  Kunde inte läsa kalendermapp för ${Mailbox}: $($_.Exception.Message)"
    }
    return $null
}

Write-Host "`nLäser gruppmedlemskap..." -ForegroundColor Cyan
$viewers = Get-GroupMemberSmtp -GroupIdentity $SeGroup
$shared  = Get-GroupMemberSmtp -GroupIdentity $DeladGroup

Write-Host "  Se-gruppen ($SeGroup): $($viewers.Count) tittare" -ForegroundColor Gray
Write-Host "  Delad-gruppen ($DeladGroup): $($shared.Count) kalendrar" -ForegroundColor Gray

if ($viewers.Count -eq 0) { Write-Warning "Inga tittare i Se-gruppen. Inget att göra."; Disconnect-ExchangeOnline -Confirm:$false; return }
if ($shared.Count  -eq 0) { Write-Warning "Inga kalendrar i Delad-gruppen. Inget att göra."; Disconnect-ExchangeOnline -Confirm:$false; return }

# Cacha kalendersökvägar för alla delade postlådor
Write-Host "`nDetekterar kalendermappar..." -ForegroundColor Cyan
$calPathCache = @{}
foreach ($mb in $shared) {
    $p = Get-CalendarFolderPath -Mailbox $mb
    if ($p) { $calPathCache[$mb] = $p; Write-Host "  $mb -> $($p.Split('\')[-1])" -ForegroundColor DarkGray }
    else    { Write-Warning "  Hoppar över $mb (ingen kalendermapp hittad)" }
}

$added = 0; $updated = 0; $removed = 0; $skipped = 0

# --- För varje delad kalender: reconcile mot önskad tittarlista ---
Write-Host "`nSynkroniserar behörigheter..." -ForegroundColor Cyan
foreach ($mb in $shared) {
    if (-not $calPathCache.ContainsKey($mb)) { continue }
    $calPath = $calPathCache[$mb]

    # Önskade tittare för DENNA kalender = alla viewers utom ägaren själv
    $wantedViewers = $viewers | Where-Object { $_ -ne $mb }

    try {
        $existing = Get-MailboxFolderPermission -Identity $calPath -ErrorAction Stop
    }
    catch {
        Write-Warning "  Kunde inte läsa behörigheter för ${mb}: $($_.Exception.Message)"
        continue
    }

    $wantedSet = @{}
    foreach ($v in $wantedViewers) { $wantedSet[$v.ToLower()] = $true }

    # 1) LÄGG TILL / UPPDATERA för varje önskad tittare
    foreach ($v in $wantedViewers) {
        $match = $existing | Where-Object {
            $_.User.ADRecipient.PrimarySmtpAddress -and
            ([string]$_.User.ADRecipient.PrimarySmtpAddress).ToLower() -eq $v.ToLower()
        } | Select-Object -First 1

        if ($null -eq $match) {
            if ($PSCmdlet.ShouldProcess("$mb : $v", "Add-MailboxFolderPermission ($PermissionLevel)")) {
                try {
                    Add-MailboxFolderPermission -Identity $calPath -User $v -AccessRights $PermissionLevel -ErrorAction Stop | Out-Null
                    Write-Host "  + $mb <- $v ($PermissionLevel)" -ForegroundColor Green
                    $added++
                }
                catch { Write-Warning "  Add misslyckades ($mb <- $v): $($_.Exception.Message)" }
            }
        }
        elseif ($match.AccessRights -notcontains $PermissionLevel) {
            # Finns men fel nivå - uppdatera endast om nuvarande är en delningsnivå vi hanterar
            # (försiktig modell: rör inte manuellt satta högre nivåer som Editor/Owner)
            if ($match.AccessRights -contains 'LimitedDetails' -or $match.AccessRights -contains 'AvailabilityOnly' -or $match.AccessRights -contains 'Reviewer') {
                if ($PSCmdlet.ShouldProcess("$mb : $v", "Set-MailboxFolderPermission -> $PermissionLevel")) {
                    try {
                        Set-MailboxFolderPermission -Identity $calPath -User $v -AccessRights $PermissionLevel -ErrorAction Stop | Out-Null
                        Write-Host "  ~ $mb <- $v (uppdaterad till $PermissionLevel)" -ForegroundColor Yellow
                        $updated++
                    }
                    catch { Write-Warning "  Set misslyckades ($mb <- $v): $($_.Exception.Message)" }
                }
            } else {
                Write-Host "  = $mb <- $v (manuell högre nivå '$($match.AccessRights -join ",")' - rörs ej)" -ForegroundColor DarkGray
                $skipped++
            }
        }
        else {
            $skipped++  # Redan rätt nivå
        }
    }

    # 2) RECONCILE: ta bort LimitedDetails-poster för tittare som INTE längre ska ha åtkomst.
    #    Försiktig modell: rör ENDAST poster med exakt behörighetsnivå = $PermissionLevel.
    #    Manuellt satta Reviewer/Editor/Owner lämnas orörda.
    foreach ($perm in $existing) {
        $permSmtp = $null
        if ($perm.User.ADRecipient.PrimarySmtpAddress) {
            $permSmtp = ([string]$perm.User.ADRecipient.PrimarySmtpAddress).ToLower()
        }
        if (-not $permSmtp) { continue }               # Default/Anonymous/okänd
        if ($permSmtp -eq $mb.ToLower()) { continue }  # ägaren själv

        $isWanted = $wantedSet.ContainsKey($permSmtp)
        $isOurLevel = ($perm.AccessRights.Count -eq 1 -and $perm.AccessRights[0] -eq $PermissionLevel)

        if ((-not $isWanted) -and $isOurLevel) {
            if ($PSCmdlet.ShouldProcess("$mb : $permSmtp", "Remove-MailboxFolderPermission (ej längre i Se-gruppen)")) {
                try {
                    Remove-MailboxFolderPermission -Identity $calPath -User $permSmtp -Confirm:$false -ErrorAction Stop
                    Write-Host "  - $mb x $permSmtp (borttagen - ej längre tittare)" -ForegroundColor Magenta
                    $removed++
                }
                catch { Write-Warning "  Remove misslyckades ($mb x $permSmtp): $($_.Exception.Message)" }
            }
        }
    }
}

Write-Host "`n===== SAMMANFATTNING =====" -ForegroundColor Cyan
Write-Host "  Tillagda:    $added" -ForegroundColor Green
Write-Host "  Uppdaterade: $updated" -ForegroundColor Yellow
Write-Host "  Borttagna:   $removed" -ForegroundColor Magenta
Write-Host "  Oförändrade: $skipped" -ForegroundColor Gray

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "`nKLART." -ForegroundColor Cyan
