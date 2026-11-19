#Requires -Version 5.1
<#
.SYNOPSIS
    Hastighetstest ned/upp PER NATVERKSKORT, ett kort i taget, en gang. Returnerar ett
    objekt per kort och kan skriva CSV.

.DESCRIPTION
    Varje aktivt kort med IPv4-adress och egen default gateway (filtrera med -AdapterName)
    testas mot en HTTPS-server med Cloudflare-kompatibla andpunkter (standard
    speed.cloudflare.com). TCP-anslutningarna binds till kortets IP-adress, sa varje kort
    mats over sin egen lank. Korten testas ett i taget eftersom de ofta delar samma
    internetlinje. For Wi-Fi-kort laggs signal % och BSSID till.

    Obs: varje test forbrukar data, ca linjehastighet x SpeedTestSec per riktning och kort.
    Cloudflare svarar 429 Too Many Requests om man kor manga tester tatt.

    Fristaende verktyg, och byggsten for:
      Network/Test-WifiConnection.ps1 -SpeedTestOnly   (anropar det har skriptet)
    Watch-ConnectionQuality.ps1 och Test-WifiConnection.ps1 kor samma motor i bakgrunden
    medan de mater latens (NetProbe.Engine.ps1).

.PARAMETER AdapterName
    Ett eller flera adapternamn, wildcards tillatna. Standard '*' = alla aktiva kort med
    IPv4-adress och default gateway.

.PARAMETER SpeedTestSec
    Matningstid per riktning (ned/upp) och kort, i sekunder. Standard 8.

.PARAMETER SpeedTestStreams
    Antal parallella TCP-anslutningar per matning. Standard 4.

.PARAMETER SpeedTestHost
    HTTPS-server med /__down?bytes=N och /__up. Standard speed.cloudflare.com.

.PARAMETER CsvPath
    Skriv resultatet till den har CSV-filen (UTF-8 med BOM). Utelamnas = ingen fil.

.PARAMETER Delimiter
    Kolumnavgransare i CSV. Standard: Windows listavgransare (';' pa svensk Windows).

.PARAMETER Quiet
    Ingen konsolutskrift, bara objekten.

.OUTPUTS
    PSCustomObject per kort: Timestamp, Computer, Adapter, Type, SourceIP, LinkSpeed,
    Server, DownMbps, UpMbps, DownMB, UpMB, Signal, BSSID, Status (OK/ERROR), Error.

.EXAMPLE
    .\Invoke-NicSpeedTest.ps1
    Alla aktiva kort, ca 16 s per kort.

.EXAMPLE
    .\Invoke-NicSpeedTest.ps1 -AdapterName 'Wi-Fi*' -SpeedTestSec 15 -CsvPath .\speed.csv

.EXAMPLE
    .\Invoke-NicSpeedTest.ps1 -Quiet | Where-Object Status -eq 'OK' | Select-Object Adapter,DownMbps,UpMbps

.NOTES
    Version : 1.0.0
    Kraver  : Windows PowerShell 5.1 eller PowerShell 7, Windows 8 / Server 2012 eller senare.
              Admin-rattigheter behovs inte. Utgaende HTTPS (443).

    Changelog:
      1.0.0  2026-11-19  Forsta version. Bruten ur Test-WifiConnection -SpeedTestOnly.
#>
[CmdletBinding()]
param(
    [string[]]$AdapterName = @('*'),
    [ValidateRange(3, 60)][int]$SpeedTestSec = 8,
    [ValidateRange(1, 16)][int]$SpeedTestStreams = 4,
    [string]$SpeedTestHost = 'speed.cloudflare.com',
    [string]$CsvPath,
    [string]$Delimiter,
    [switch]$Quiet
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0.0'

$engine = Join-Path $PSScriptRoot 'NetProbe.Engine.ps1'
if (-not (Test-Path -LiteralPath $engine)) { throw "NetProbe.Engine.ps1 saknas bredvid $($MyInvocation.MyCommand.Name) - kopiera hela Network-mappen." }
. $engine

function Out-Line([string]$Text, [string]$Color = 'Gray') { if (-not $Quiet) { Write-Host $Text -ForegroundColor $Color } }
function Format-Mbps([double]$v) { if ($v -lt 0) { 'fel' } else { [math]::Round($v, 1).ToString([cultureinfo]::InvariantCulture) } }

$all  = @(Get-NetProbeNic -Name $AdapterName)
foreach ($s in @($all | Where-Object SkipReason)) { Out-Line "Hoppar over '$($s.Name)': $($s.SkipReason)." 'DarkYellow' }
$nics = @($all | Where-Object { -not $_.SkipReason })
if ($nics.Count -eq 0) {
    throw "Inga aktiva kort med IPv4 och default gateway matchar '$($AdapterName -join "', '")'. Kontrollera med: Get-NetAdapter | Where-Object Status -eq 'Up'"
}

# Wi-Fi-signal per kort (fore testet)
$wlan = @{}
if (@($nics | Where-Object IsWlan).Count -gt 0) { $wlan = (Get-WlanSnapshot).Map }

Out-Line ("Invoke-NicSpeedTest v{0}: {1}, {2} s ned + {2} s upp per kort, {3} strommar, {4} kort ett i taget (ca {5} s)" -f `
    $ScriptVersion, $SpeedTestHost, $SpeedTestSec, $SpeedTestStreams, $nics.Count, ($nics.Count * 2 * ($SpeedTestSec + 1))) 'Cyan'

$results = $null
$task = $null
try {
    $task = [SpeedTestV1]::StartAll([string[]]@($nics | ForEach-Object { $_.Src }), $SpeedTestHost, $SpeedTestSec, $SpeedTestStreams)
    $shown = -1
    while (-not $task.IsCompleted) {
        $cur = [SpeedTestV1]::Current
        if ($cur -ge 0 -and $cur -ne $shown) {
            Out-Line ("{0}  Testar [{1}] fran {2} ..." -f (Get-Date -Format 'HH:mm:ss'), $nics[$cur].Name, $nics[$cur].Src) 'DarkCyan'
            $shown = $cur
        }
        Start-Sleep -Milliseconds 200
    }
    $results = $task.Result
} finally {
    if (-not $results) {
        [SpeedTestV1]::Cancel()
        Out-Line 'AVBRUTET - inga resultat.' 'Yellow'
    }
}

$rows = foreach ($r in $results) {
    $n = $nics[$r.Index]
    $w = if ($n.IsWlan) { $wlan[$n.Guid] } else { $null }
    [PSCustomObject]@{
        Timestamp = $r.Started.ToString('yyyy-MM-dd HH:mm:ss')
        Computer  = $env:COMPUTERNAME
        Adapter   = $n.Name
        Type      = $n.Typ
        SourceIP  = $n.Src
        LinkSpeed = $n.LinkSpeed
        Server    = $SpeedTestHost
        DownMbps  = if ($r.DownMbps -ge 0) { [math]::Round($r.DownMbps, 1) } else { $null }
        UpMbps    = if ($r.UpMbps -ge 0) { [math]::Round($r.UpMbps, 1) } else { $null }
        DownMB    = [math]::Round($r.DownBytes / 1MB, 1)
        UpMB      = [math]::Round($r.UpBytes / 1MB, 1)
        Signal    = if ($w) { $w.Signal } else { $null }
        BSSID     = if ($w) { $w.BSSID } else { $null }
        Status    = if ($r.Error) { 'ERROR' } else { 'OK' }
        Error     = if ($r.Error) { $r.Error } else { $null }
    }
}
$rows = @($rows)

foreach ($x in $rows) {
    $wtxt = if ($null -ne $x.Signal) { "  signal=$($x.Signal)%  BSSID=$($x.BSSID)" } else { '' }
    $etxt = if ($x.Error) { "  ($($x.Error))" } else { '' }
    $down = if ($null -ne $x.DownMbps) { $x.DownMbps.ToString([cultureinfo]::InvariantCulture) } else { 'fel' }
    $up   = if ($null -ne $x.UpMbps)   { $x.UpMbps.ToString([cultureinfo]::InvariantCulture) }   else { 'fel' }
    Out-Line ("{0}  [{1}] ned {2} Mbit/s  upp {3} Mbit/s  ({4} + {5} MB){6}{7}" -f `
        $x.Timestamp.Substring(11), $x.Adapter, $down, $up, $x.DownMB, $x.UpMB, $wtxt, $etxt) $(if ($x.Error) { 'Yellow' } else { 'Green' })
}

if ($CsvPath -and $rows.Count) {
    if (-not $Delimiter) { $Delimiter = (Get-Culture).TextInfo.ListSeparator }
    if (-not $Delimiter) { $Delimiter = ',' }
    $CsvPath = [System.IO.Path]::GetFullPath($CsvPath)
    # Export-Csv on PS 5.1 has no BOM option for UTF8 that Excel likes consistently - write it ourselves.
    $lines = @($rows | ConvertTo-Csv -NoTypeInformation -Delimiter $Delimiter)
    [System.IO.File]::WriteAllLines($CsvPath, [string[]]$lines, (New-Object System.Text.UTF8Encoding($true)))
    Out-Line "Sparat: $CsvPath"
}

$rows
