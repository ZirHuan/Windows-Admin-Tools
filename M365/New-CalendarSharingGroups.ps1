#Requires -Version 7.0
<#
.SYNOPSIS
    Skapar de två statiska säkerhetsgrupperna för kalenderdelning och fyller dem.

.DESCRIPTION
    Skapar:
      - Kal-Delad-Antilop : postlådor vars kalendrar exponeras (10 medlemmar)
      - Kal-Se-Antilop    : tittare (endast Carl Mönefors initialt, för funktionstest)

    Grupperna skapas som mail-enabled security groups (distributionsgrupp med
    säkerhetsbehörighet) eftersom Sync-CalendarSharing.ps1 läser medlemmar via
    Get-DistributionGroupMember. Rena Entra-säkerhetsgrupper (icke mail-enabled)
    kan INTE läsas med det cmdletet - därför mail-enabled security här.

    Medlemslistan är verifierad mot users.json (audit 2026-08-10, mapp 20260810_1011):
    endast aktiverade, licensierade personkonton. Exkluderade: DMARC-konton,
    iver-admin, tjänstekonton (svc.m365backup), Toridium-konton (disabled), gäster.

    OBS gabriel@antilopgroup.se: dubblettkonto med annan licens, medtaget enligt
    beslut (accepterad risk - historik ej maskerad, kontot under utredning).

.PARAMETER WhatIf
    Torrkörning - visar vad som skulle skapas utan att skapa något.

.NOTES
    Kräver: ExchangeOnlineManagement-modulen, behörighet att skapa distributionsgrupper.
    KÖR MED -WhatIf FÖRST.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$AppOnly,
    [string]$AppId,
    [string]$CertThumbprint,
    [string]$Organization
)

$ScriptVersion = '1.1.0'
Write-Host "New-CalendarSharingGroups v$ScriptVersion" -ForegroundColor Cyan

# --- Gruppdefinitioner ---
$DeladName  = 'Kal-Delad-Antilop'
$DeladAlias = 'kal-delad-antilop'
$SeName     = 'Kal-Se-Antilop'
$SeAlias    = 'kal-se-antilop'

# Verifierade medlemmar (users.json, audit 20260810_1011)
$DeladMembers = @(
    'alma.bratt@antilopgroup.se'
    'carl.monefors@antilopgroup.se'
    'filip.westin@antilopgroup.se'
    'gabriel.sjostrom@antilopgroup.se'
    'gabriel@antilopgroup.se'          # dubblett, accepterad risk
    'henrik.prior@restra.se'
    'joakim.blom@restra.se'
    'kent.bjork@antilopgroup.se'
    'marvin.stenersen@antilopgroup.se'
    'william.jaeckel@antilopgroup.se'
)

$SeMembers = @(
    'carl.monefors@antilopgroup.se'
    'philip.vitzthum@antilopgroup.se'
)

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

# --- Hjälpfunktion: skapa grupp om den inte finns, fyll medlemmar ---
function New-CalGroup {
    param(
        [string]$Name,
        [string]$Alias,
        [string[]]$Members
    )

    $existing = Get-DistributionGroup -Identity $Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  Gruppen '$Name' finns redan - hoppar över skapande, synkar medlemmar." -ForegroundColor Yellow
    }
    else {
        if ($PSCmdlet.ShouldProcess($Name, "New-DistributionGroup (mail-enabled security)")) {
            try {
                New-DistributionGroup -Name $Name -Alias $Alias -Type 'Security' `
                    -PrimarySmtpAddress "$Alias@antilopgroup.se" -ErrorAction Stop | Out-Null
                Write-Host "  + Skapade grupp '$Name'" -ForegroundColor Green
                # Vänta så gruppen hinner propagera innan medlemmar läggs till
                Start-Sleep -Seconds 5
            }
            catch { throw "Kunde inte skapa '$Name': $($_.Exception.Message)" }
        }
    }

    # Lägg till medlemmar (idempotent - hoppar över befintliga)
    foreach ($m in $Members) {
        $already = $null
        try {
            $already = Get-DistributionGroupMember -Identity $Name -ResultSize Unlimited -ErrorAction SilentlyContinue |
                       Where-Object { ([string]$_.PrimarySmtpAddress).ToLower() -eq $m.ToLower() }
        } catch { }

        if ($already) {
            Write-Host "    = $m (redan medlem)" -ForegroundColor DarkGray
            continue
        }
        if ($PSCmdlet.ShouldProcess("$Name : $m", "Add-DistributionGroupMember")) {
            try {
                Add-DistributionGroupMember -Identity $Name -Member $m -ErrorAction Stop
                Write-Host "    + $m" -ForegroundColor Green
            }
            catch { Write-Warning "    Kunde inte lägga till $m i ${Name}: $($_.Exception.Message)" }
        }
    }
}

Write-Host "`nSkapar Delad-gruppen..." -ForegroundColor Cyan
New-CalGroup -Name $DeladName -Alias $DeladAlias -Members $DeladMembers

Write-Host "`nSkapar Se-gruppen..." -ForegroundColor Cyan
New-CalGroup -Name $SeName -Alias $SeAlias -Members $SeMembers

Write-Host "`nKLART. Grupperna är skapade/uppdaterade." -ForegroundColor Cyan
Write-Host "Nästa steg: kör Sync-CalendarSharing.ps1 med -WhatIf" -ForegroundColor Cyan

Disconnect-ExchangeOnline -Confirm:$false | Out-Null
