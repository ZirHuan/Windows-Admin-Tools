<#
.SYNOPSIS
    Loggar latens mot gateway och internet per natverkskort (alla aktiva NIC), Wi-Fi-
    signal/BSSID for tradlosa kort och gor periodiska hastighetstester per kort, for att
    hitta orsaken till latenstoppar och lag hastighet.

.DESCRIPTION
    LATENS: Varje intervall (standard 1 s) skickas ICMP echo PARALLELLT fran varje NIC:s
    egen IPv4-adress till kortets default gateway och till ett internetmal. Pingen gors
    med Windows IcmpSendEcho2Ex (iphlpapi.dll) med kalladressen bunden - samma sak som
    'ping -S', men utan att starta ping.exe och utan sprakberoende texttolkning.

    WI-FI: For tradlosa kort laggs SSID, BSSID, signal %, RSSI (om Windows visar det),
    kanal, radiotyp och rx/tx-hastighet till fran 'netsh wlan show interfaces'. Blocken
    kopplas till ratt kort via GUID (flera Wi-Fi-kort, svensk/engelsk Windows).

    HASTIGHET: Med jamna mellanrum (standard: 10 s efter start och sedan var 180:e minut)
    mats nedladdning och uppladdning PER KORT mot ett HTTPS-hastighetstest (standard
    speed.cloudflare.com). TCP-anslutningarna binds till kortets IP, sa varje kort mats
    over sin egen lank. Korten mats ett i taget (de delar ofta samma internetlinje).
    Testet kors i bakgrunden; latensmatningen fortsatter under tiden men de raderna
    markeras 'UnderLast' och raknas inte som toppar - de redovisas separat som latens
    under last (bufferbloat). Obs: varje test forbrukar data (ca linjehastighet x
    SpeedTestSec per riktning och kort).

    Matningarna schemalaggs pa fasta tidsluckor (ingen drift). Resultaten skrivs lopande
    till CSV (latens: en rad per NIC och matning; hastighet: egen fil) och overlever
    Ctrl+C. Sammanfattningen visar per NIC forlust, median/p95/max, toppar, roaming,
    hastighet och om topparna kommer samtidigt pa flera kort eller bara pa ett.

.PARAMETER AdapterName
    Ett eller flera adapternamn, wildcards tillatna. Standard '*' = alla aktiva kort
    med IPv4-adress och default gateway. Exempel: -AdapterName 'Wi-Fi*','Ethernet'

.PARAMETER DurationSec
    Testets langd i sekunder. Standard 300. Ett dygn = 86400.

.PARAMETER IntervalMs
    Tid mellan latensmatningar i ms. Standard 1000.

.PARAMETER TimeoutMs
    Svarstid innan en ping raknas som timeout. Maste vara mindre an IntervalMs.
    Standard 800.

.PARAMETER SpikeMs
    Gateway-latens over detta varde (ms) raknas som topp. Standard 30.

.PARAMETER InternetSpikeMs
    Internet-latens over detta varde (ms) raknas som topp. Standard 0 = 2 x SpikeMs.

.PARAMETER InternetTarget
    IPv4-adress eller vardnamn for latensmatningen mot internet. Standard 8.8.8.8.

.PARAMETER RoamWindowSec
    En topp inom sa manga sekunder fran ett BSSID-byte raknas som "nara roaming".
    Standard 2.

.PARAMETER SpeedTestIntervalMin
    Minuter mellan hastighetstester. Forsta testet kors 10 s efter start.
    Standard 180 (var tredje timme). 0 = inga hastighetstester.

.PARAMETER SpeedTestOnly
    Bara hastighetstest: anropar Invoke-NicSpeedTest.ps1 (samma mapp), som testar varje
    valt kort EN gang (ned + upp), och avslutar. Ingen latensmatning.
    -DurationSec, -IntervalMs och -SpeedTestIntervalMin ignoreras.
    Resultatet sparas i -LogPath (standard Skrivbordet\netspeed-<datum>-<tid>.csv).
    Kor hellre Invoke-NicSpeedTest.ps1 direkt - det returnerar objekt.

.PARAMETER SpeedTestSec
    Matningstid per riktning (ned/upp) och kort, i sekunder. Standard 8.

.PARAMETER SpeedTestStreams
    Antal parallella TCP-anslutningar per matning. Standard 4.

.PARAMETER SpeedTestHost
    HTTPS-server med Cloudflare-kompatibla andpunkter /__down?bytes=N och /__up.
    Standard speed.cloudflare.com. Obs: Cloudflare svarar 403 pa >= 100 MB per
    begaran (skriptet hamtar 50 MB at gangen) och 429 Too Many Requests om man
    kor manga tester tatt - sett efter ca 20 tester inom 1,5 h. Ett 429 loggas som
    fel for det testet; latensmatningen fortsatter.

.PARAMETER LogPath
    CSV-fil for latens. Standard: Skrivbordet\netconn-<datum>-<tid>.csv (TEMP om
    Skrivbord saknas). Hastighetstesterna sparas bredvid som <namn>-hastighet.csv.

.PARAMETER Delimiter
    Kolumnavgransare i CSV. Standard: Windows listavgransare (';' pa svensk Windows,
    sa filen oppnas ratt direkt i Excel).

.EXAMPLE
    .\Test-WifiConnection.ps1
    Alla aktiva kort i 5 minuter, ett hastighetstest per kort efter 10 s.

.EXAMPLE
    .\Test-WifiConnection.ps1 -DurationSec 86400
    Ett dygn, hastighetstest per kort var tredje timme (9 tester per kort).

.EXAMPLE
    .\Test-WifiConnection.ps1 -AdapterName 'Wi-Fi' -DurationSec 900 -SpikeMs 25 -SpeedTestIntervalMin 0
    Bara Wi-Fi-kortet i 15 minuter, topp-grans 25 ms, inga hastighetstester.

.EXAMPLE
    .\Test-WifiConnection.ps1 -SpeedTestOnly
    Hastighetstest ned/upp pa alla aktiva kort, ett i taget, och avsluta (ca 16 s per kort).

.EXAMPLE
    .\Test-WifiConnection.ps1 -SpeedTestOnly -AdapterName 'Wi-Fi*' -SpeedTestSec 15
    Bara Wi-Fi-korten, 15 s per riktning.

.NOTES
    Version : 1.2.0
    Kraver  : Windows PowerShell 5.1 eller PowerShell 7, Windows 8 / Server 2012 eller
              senare. Admin-rattigheter behovs inte. Utgaende HTTPS (443) for
              hastighetstestet.

    Byggstenar (samma mapp): NetProbe.Engine.ps1 (ICMP, hastighetstest, kortval,
    netsh wlan) och Invoke-NicSpeedTest.ps1 (-SpeedTestOnly).

    Changelog
      1.2.0  2026-11-19  C#-motorn, kortvalet och netsh-tolkningen flyttade till
                         NetProbe.Engine.ps1 (delas med Watch-ConnectionQuality).
                         -SpeedTestOnly anropar nya Invoke-NicSpeedTest.ps1; dess CSV
                         har nu engelska kolumnnamn (samma som Watch-ConnectionQuality).
      1.1.0  2026-11-04  -SpeedTestOnly: bara hastighetstest per kort, en gang, sedan
                         avslut. Ingen latensmatning.
      1.0.0  2026-11-03  Ombyggd fran forsta versionen:
                         - Alla aktiva NIC testas samtidigt (-AdapterName filtrerar).
                         - Parallell ICMP via IcmpSendEcho2Ex med bunden kalladress
                           i stallet for seriella ping.exe-anrop och textparsning.
                         - Fasta tidsluckor, tidsstampel vid matningens start,
                           roamingfonster i sekunder, overskridningar raknas.
                         - netsh-block kopplas till kort via GUID (flera Wi-Fi-kort).
                         - Gateway fran kortets egen default route (lagst metric).
                         - Periodiskt hastighetstest ned/upp per kort (bunden kalla),
                           latens under last redovisas separat.
                         - CSV skrivs lopande, Excel-vanlig avgransare.
                         - Statistik per NIC (median/p95/max/forlust) och
                           korrelation av toppar mellan NIC.
#>
[CmdletBinding()]
param(
    [string[]]$AdapterName = @('*'),
    [ValidateRange(5, 604800)][int]$DurationSec = 300,
    [ValidateRange(200, 60000)][int]$IntervalMs = 1000,
    [ValidateRange(50, 30000)][int]$TimeoutMs = 800,
    [ValidateRange(1, 10000)][int]$SpikeMs = 30,
    [ValidateRange(0, 30000)][int]$InternetSpikeMs = 0,
    [string]$InternetTarget = '8.8.8.8',
    [ValidateRange(0, 60)][int]$RoamWindowSec = 2,
    [ValidateRange(0, 10080)][int]$SpeedTestIntervalMin = 180,
    [switch]$SpeedTestOnly,
    [ValidateRange(3, 60)][int]$SpeedTestSec = 8,
    [ValidateRange(1, 16)][int]$SpeedTestStreams = 4,
    [string]$SpeedTestHost = 'speed.cloudflare.com',
    [string]$LogPath,
    [string]$Delimiter
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# -SpeedTestOnly ar ett eget verktyg sedan 1.2.0 - skicka vidare och avsluta.
if ($SpeedTestOnly) {
    $tool = Join-Path $PSScriptRoot 'Invoke-NicSpeedTest.ps1'
    if (-not (Test-Path -LiteralPath $tool)) { throw "Invoke-NicSpeedTest.ps1 saknas bredvid Test-WifiConnection.ps1 - kopiera hela Network-mappen." }
    if (-not $LogPath) {
        $dir = [Environment]::GetFolderPath('Desktop')
        if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { $dir = $env:TEMP }
        $LogPath = Join-Path $dir ("netspeed-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    }
    $fwd = @{ AdapterName = $AdapterName; SpeedTestSec = $SpeedTestSec; SpeedTestStreams = $SpeedTestStreams
              SpeedTestHost = $SpeedTestHost; CsvPath = $LogPath }
    if ($Delimiter) { $fwd.Delimiter = $Delimiter }
    & $tool @fwd | Format-Table Adapter, Type, SourceIP, LinkSpeed, DownMbps, UpMbps, Signal, Status -AutoSize | Out-Host
    return
}

# ---------- Delad motor: ICMP, hastighetstest, kortval, netsh wlan ----------
$engine = Join-Path $PSScriptRoot 'NetProbe.Engine.ps1'
if (-not (Test-Path -LiteralPath $engine)) { throw "NetProbe.Engine.ps1 saknas bredvid Test-WifiConnection.ps1 - kopiera hela Network-mappen." }
. $engine

# ---------- Hjalpfunktioner ----------
function Get-Percentile($Values, [double]$P) {
    if ($null -eq $Values -or $Values.Count -eq 0) { return $null }
    $a = @($Values | Sort-Object)
    $idx = [int][math]::Ceiling($P * $a.Length) - 1
    if ($idx -lt 0) { $idx = 0 }
    return $a[$idx]
}

function Format-Stats($Values) {
    if ($null -eq $Values -or $Values.Count -eq 0) { return '-' }
    return '{0}/{1}/{2}' -f (Get-Percentile $Values 0.5), (Get-Percentile $Values 0.95), (Get-Percentile $Values 1.0)
}

function Format-Rtt([int]$raw) {
    if ($raw -ge 0) { return "$raw ms" }
    if ($raw -eq -11010) { return 'timeout' }
    return "fel $(-$raw)"
}

function Format-Mbps([double]$v) {
    if ($v -lt 0) { return 'fel' }
    return [math]::Round($v, 1).ToString([cultureinfo]::InvariantCulture)
}

function Write-CsvRow($Writer, [object[]]$Values) {
    $out = foreach ($v in $Values) {
        if ($null -eq $v) { ''; continue }
        $s = [string]$v
        if ($s.IndexOf($delimChar) -ge 0 -or $s.IndexOf('"') -ge 0 -or $s.IndexOf("`n") -ge 0) {
            '"' + $s.Replace('"', '""') + '"'
        } else { $s }
    }
    $Writer.WriteLine(($out -join $Delimiter))
}

function New-StatSet {
    [pscustomobject]@{
        N      = 0
        GwOk   = [System.Collections.Generic.List[int]]::new()
        InOk   = [System.Collections.Generic.List[int]]::new()
        GwLost = 0
        InLost = 0
    }
}

function Complete-SpeedTest($Results) {
    foreach ($r in $Results) {
        $n = $nics[$r.Index]
        $sig = if ($n.LastWlan) { $n.LastWlan.Signal } else { $null }
        $bss = if ($n.LastWlan) { $n.LastWlan.BSSID } else { $null }
        $n.Speeds.Add($r)
        $col = if ($r.Error) { 'Yellow' } else { 'Green' }
        $wtxt = if ($null -ne $sig) { "  signal=$sig%  BSSID=$bss" } else { '' }
        $etxt = if ($r.Error) { "  ($($r.Error))" } else { '' }
        Write-Host ("{0}  HASTIGHET [{1}]: ned {2} Mbit/s  upp {3} Mbit/s  ({4:N0} + {5:N0} MB){6}{7}" -f `
            $r.Started.ToString('HH:mm:ss'), $n.Name, (Format-Mbps $r.DownMbps), (Format-Mbps $r.UpMbps),
            ($r.DownBytes / 1MB), ($r.UpBytes / 1MB), $wtxt, $etxt) -ForegroundColor $col
        Write-CsvRow $speedWriter @(
            $r.Started.ToString('yyyy-MM-dd HH:mm:ss'), $n.Name, $n.Typ, $n.Src, $SpeedTestHost,
            $(if ($r.DownMbps -ge 0) { Format-Mbps $r.DownMbps }), $(if ($r.UpMbps -ge 0) { Format-Mbps $r.UpMbps }),
            [math]::Round($r.DownBytes / 1MB, 1).ToString([cultureinfo]::InvariantCulture),
            [math]::Round($r.UpBytes / 1MB, 1).ToString([cultureinfo]::InvariantCulture),
            $sig, $bss, $(if ($r.Error) { $r.Error } else { 'ok' })
        )
    }
    $speedWriter.Flush()
}

# ---------- Forberedelser ----------
if ($TimeoutMs -ge $IntervalMs) {
    Write-Host "TimeoutMs ($TimeoutMs) maste vara mindre an IntervalMs ($IntervalMs)." -ForegroundColor Red
    exit 1
}
if ($InternetSpikeMs -le 0) { $InternetSpikeMs = 2 * $SpikeMs }
if (-not $Delimiter) { $Delimiter = (Get-Culture).TextInfo.ListSeparator }
if (-not $Delimiter) { $Delimiter = ',' }
$delimChar = $Delimiter[0]
$speedOn = $SpeedTestIntervalMin -gt 0

# Internetmal -> IPv4
$inetIP = $null
$parsed = $null
if ([System.Net.IPAddress]::TryParse($InternetTarget, [ref]$parsed) -and
    $parsed.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
    $inetIP = $parsed.ToString()
} else {
    try {
        $inetIP = [System.Net.Dns]::GetHostAddresses($InternetTarget) |
            Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } |
            Select-Object -First 1 | ForEach-Object { $_.ToString() }
    } catch { $inetIP = $null }
}
if (-not $inetIP) {
    Write-Host "Kunde inte fa fram en IPv4-adress for internetmalet '$InternetTarget'." -ForegroundColor Red
    exit 1
}

# Kort (urvalet gors i NetProbe.Engine.ps1, samma som i Watch-ConnectionQuality)
$nics = [System.Collections.Generic.List[object]]::new()
foreach ($c in @(Get-NetProbeNic -Name $AdapterName)) {
    if ($c.SkipReason) { Write-Host "Hoppar over '$($c.Name)': $($c.SkipReason)." -ForegroundColor DarkYellow; continue }
    $nics.Add([pscustomobject]@{
        Name      = $c.Name
        Guid      = $c.Guid
        Typ       = $c.Typ
        IsWlan    = $c.IsWlan
        Src       = $c.Src
        Gw        = $c.Gw
        Speed     = $c.LinkSpeed
        Idle      = New-StatSet      # latens utan hastighetstest
        Load      = New-StatSet      # latens medan hastighetstest pagar
        Spikes    = [System.Collections.Generic.List[int]]::new()   # matningsindex
        Roams     = [System.Collections.Generic.List[int]]::new()   # matningsindex
        Speeds    = [System.Collections.Generic.List[object]]::new()
        PrevBssid = $null
        LastWlan  = $null
        SigSpikeSum = 0; SigSpikeN = 0; SigNormSum = 0; SigNormN = 0
        Aps       = @{}   # BSSID -> @{N; SigSum; SigN; SigMin}
    })
}
if ($nics.Count -eq 0) {
    Write-Host "Inga aktiva kort med IPv4 och default gateway matchar '$($AdapterName -join "', '")'." -ForegroundColor Red
    Write-Host "Kontrollera med: Get-NetAdapter | Where-Object Status -eq 'Up'" -ForegroundColor Red
    exit 1
}
$anyWlan = @($nics | Where-Object IsWlan).Count -gt 0

# Loggfiler
if (-not $LogPath) {
    $dir = [Environment]::GetFolderPath('Desktop')
    if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { $dir = $env:TEMP }
    $LogPath = Join-Path $dir ("netconn-{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
}
$LogPath = [System.IO.Path]::GetFullPath($LogPath)
$SpeedLogPath = Join-Path ([System.IO.Path]::GetDirectoryName($LogPath)) `
    ([System.IO.Path]::GetFileNameWithoutExtension($LogPath) + '-hastighet.csv')
$utf8Bom = New-Object System.Text.UTF8Encoding($true)

$writer = New-Object System.IO.StreamWriter($LogPath, $false, $utf8Bom)
$writer.WriteLine((@('Tid','Matning','Adapter','Typ','KallIP','Gateway','GatewayMs','GatewayStatus',
    'InternetMs','InternetStatus','Topp','Orsak','UnderLast','SSID','BSSID','Signal','Rssi','Kanal',
    'Band','RxRate','TxRate','Roaming') -join $Delimiter))
$speedWriter = $null
if ($speedOn) {
    $speedWriter = New-Object System.IO.StreamWriter($SpeedLogPath, $false, $utf8Bom)
    $speedWriter.WriteLine((@('Tid','Adapter','Typ','KallIP','Server','NedMbps','UppMbps','NedMB','UppMB',
        'Signal','BSSID','Status') -join $Delimiter))
}

# Startinfo
Write-Host ""
Write-Host "Internetmal: $InternetTarget ($inetIP) | Intervall: $IntervalMs ms | Timeout: $TimeoutMs ms" -ForegroundColor Cyan
Write-Host "Topp-granser: gateway > $SpikeMs ms, internet > $InternetSpikeMs ms, eller timeout" -ForegroundColor Cyan
if ($speedOn) {
    Write-Host ("Hastighetstest: {0}, forst efter 10 s och sedan var {1}:e minut, {2} s ned + {2} s upp per kort, {3} strommar" -f `
        $SpeedTestHost, $SpeedTestIntervalMin, $SpeedTestSec, $SpeedTestStreams) -ForegroundColor Cyan
} else {
    Write-Host "Hastighetstest: av" -ForegroundColor Cyan
}
$nics | Format-Table @{n='Kort';e={$_.Name}}, Typ, @{n='Kalla';e={$_.Src}}, @{n='Gateway';e={$_.Gw}},
                     @{n='Lankhastighet';e={$_.Speed}} -AutoSize | Out-Host

if ($anyWlan) {
    $wlanFirst = Get-WlanSnapshot
    foreach ($n in ($nics | Where-Object IsWlan)) {
        $w = $wlanFirst.Map[$n.Guid]
        if (-not $w -or -not $w.BSSID) {
            Write-Host "VARNING: hittar inget BSSID for '$($n.Name)' i netsh. Ratt utskrift:" -ForegroundColor Yellow
            $wlanFirst.Raw | Out-Host
            Write-Host "Skicka den utskriften sa justerar vi falt-namnen. Matningen fortsatter anda." -ForegroundColor Yellow
        } else {
            Write-Host ("[{0}] SSID: {1} | BSSID: {2} | Signal: {3}% | Kanal: {4} | Radio: {5}" -f `
                $n.Name, $w.SSID, $w.BSSID, $w.Signal, $w.Kanal, $w.Radiotyp)
            $n.PrevBssid = $w.BSSID
            $n.LastWlan = $w
        }
    }
}

Write-Host "`nLoggar i $DurationSec s till $LogPath"
if ($speedOn) { Write-Host "Hastighetstester till $SpeedLogPath" }
Write-Host "Toppar, roaming och hastighet visas nedan. Avbryt med Ctrl+C (sammanfattningen skrivs anda).`n"

# ---------- Huvudloop ----------
$srcArr = [string[]]@(foreach ($n in $nics) { $n.Src; $n.Src })
$dstArr = [string[]]@(foreach ($n in $nics) { $n.Gw;  $inetIP })
$nicSrc = [string[]]@(foreach ($n in $nics) { $n.Src })
$tickTimes = [System.Collections.Generic.List[datetime]]::new()
$spikeTicks = @{}           # matningsindex -> antal NIC med topp
$overruns = 0
$tick = 0
$clock = [Diagnostics.Stopwatch]::StartNew()
$endMs = [int64]$DurationSec * 1000
$nextAt = [int64]0
$nextStatus = [int64]60000
$nextSpeedAt = [int64]10000
$speedTask = $null
$completed = $false

try {
    while ($clock.ElapsedMilliseconds -lt $endMs) {
        # hastighetstest: starta / skorda
        if ($speedOn) {
            if ($speedTask -and $speedTask.IsCompleted) {
                Complete-SpeedTest $speedTask.Result
                $speedTask = $null
            }
            # jamfor mot planerad tidslucka ($nextAt), inte klockan - Start-Sleep kan vakna nagra ms tidigt
            if (-not $speedTask -and $nextAt -ge $nextSpeedAt) {
                $speedTask = [SpeedTestV1]::StartAll($nicSrc, $SpeedTestHost, $SpeedTestSec, $SpeedTestStreams)
                Write-Host ("{0}  Hastighetstest startat ({1} kort, ca {2} s)" -f (Get-Date -Format 'HH:mm:ss'),
                    $nics.Count, ($nics.Count * 2 * $SpeedTestSec)) -ForegroundColor DarkCyan
                $nextSpeedAt += [int64]$SpeedTestIntervalMin * 60000
            }
        }
        $underLoad = $null -ne $speedTask

        $now = Get-Date
        $tickTimes.Add($now)
        $tasks = [NetConnProbeV1]::Start($srcArr, $dstArr, $TimeoutMs)
        $wl = if ($anyWlan) { (Get-WlanSnapshot).Map } else { @{} }
        $res = [NetConnProbeV1]::Collect($tasks)
        $tidText = $now.ToString('yyyy-MM-dd HH:mm:ss.fff')
        $spikeNics = 0

        for ($j = 0; $j -lt $nics.Count; $j++) {
            $n = $nics[$j]
            $gwRaw = $res[2 * $j]
            $inRaw = $res[2 * $j + 1]
            $st = if ($underLoad) { $n.Load } else { $n.Idle }
            $st.N++
            if ($gwRaw -ge 0) { $st.GwOk.Add($gwRaw) } else { $st.GwLost++ }
            if ($inRaw -ge 0) { $st.InOk.Add($inRaw) } else { $st.InLost++ }

            $why = [System.Collections.Generic.List[string]]::new()
            if ($gwRaw -lt 0) { $why.Add("gw " + (Format-Rtt $gwRaw)) } elseif ($gwRaw -gt $SpikeMs) { $why.Add("gw $gwRaw ms") }
            if ($inRaw -lt 0) { $why.Add("inet " + (Format-Rtt $inRaw)) } elseif ($inRaw -gt $InternetSpikeMs) { $why.Add("inet $inRaw ms") }
            # under hastighetstest ar hog latens vantad - logga men rakna inte som topp
            $isSpike = ($why.Count -gt 0) -and -not $underLoad

            $w = $null; $roamed = $false
            if ($n.IsWlan) {
                $w = $wl[$n.Guid]
                if ($w -and $w.BSSID) {
                    $n.LastWlan = $w
                    if ($n.PrevBssid -and $w.BSSID -ne $n.PrevBssid) {
                        $roamed = $true
                        $n.Roams.Add($tick)
                        Write-Host ("{0}  ROAMING [{1}]: {2} -> {3} (signal {4}%, kanal {5})" -f `
                            $now.ToString('HH:mm:ss'), $n.Name, $n.PrevBssid, $w.BSSID, $w.Signal, $w.Kanal) -ForegroundColor Magenta
                    }
                    $n.PrevBssid = $w.BSSID
                    if (-not $n.Aps.ContainsKey($w.BSSID)) { $n.Aps[$w.BSSID] = @{ N = 0; SigSum = 0; SigN = 0; SigMin = 101 } }
                    $ap = $n.Aps[$w.BSSID]; $ap.N++
                    if ($null -ne $w.Signal) {
                        $ap.SigSum += $w.Signal; $ap.SigN++
                        if ($w.Signal -lt $ap.SigMin) { $ap.SigMin = $w.Signal }
                        if (-not $underLoad) {
                            if ($isSpike) { $n.SigSpikeSum += $w.Signal; $n.SigSpikeN++ }
                            else          { $n.SigNormSum  += $w.Signal; $n.SigNormN++ }
                        }
                    }
                }
            }

            if ($isSpike) {
                $n.Spikes.Add($tick); $spikeNics++
                $extra = if ($w) { "  signal=$($w.Signal)%  BSSID=$($w.BSSID)  kanal=$($w.Kanal)" } else { '' }
                Write-Host ("{0}  TOPP [{1}]: gw={2}  inet={3}{4}" -f $now.ToString('HH:mm:ss'), $n.Name,
                    (Format-Rtt $gwRaw), (Format-Rtt $inRaw), $extra) -ForegroundColor Red
            }

            Write-CsvRow $writer @(
                $tidText, $tick, $n.Name, $n.Typ, $n.Src, $n.Gw,
                $(if ($gwRaw -ge 0) { $gwRaw } else { $null }), $(if ($gwRaw -ge 0) { 'ok' } else { Format-Rtt $gwRaw }),
                $(if ($inRaw -ge 0) { $inRaw } else { $null }), $(if ($inRaw -ge 0) { 'ok' } else { Format-Rtt $inRaw }),
                $isSpike, ($why -join '; '), $underLoad,
                $(if ($w) { $w.SSID }), $(if ($w) { $w.BSSID }), $(if ($w) { $w.Signal }), $(if ($w) { $w.Rssi }),
                $(if ($w) { $w.Kanal }), $(if ($w) { $w.Band }), $(if ($w) { $w.RxRate }), $(if ($w) { $w.TxRate }),
                $roamed
            )
        }
        $writer.Flush()
        if ($spikeNics -gt 0) { $spikeTicks[$tick] = $spikeNics }
        $tick++

        if ($clock.ElapsedMilliseconds -ge $nextStatus) {
            Write-Host ("{0}  ... {1}/{2} s, {3} matningar, {4} med topp" -f (Get-Date -Format 'HH:mm:ss'),
                [int]($clock.ElapsedMilliseconds / 1000), $DurationSec, $tick, $spikeTicks.Count) -ForegroundColor DarkGray
            $nextStatus += 60000
        }

        $nextAt += $IntervalMs
        $wait = $nextAt - $clock.ElapsedMilliseconds
        if ($wait -gt 0) {
            if ($nextAt -ge $endMs) { break }
            Start-Sleep -Milliseconds ([int]$wait)
        } else {
            $overruns++
            $nextAt = $clock.ElapsedMilliseconds
        }
    }

    # ett pagaende hastighetstest far bli klart (det ar tidsbegransat)
    if ($speedTask) {
        Write-Host "Vantar pa pagaende hastighetstest..." -ForegroundColor DarkCyan
        while (-not $speedTask.IsCompleted) { Start-Sleep -Milliseconds 200 }
        Complete-SpeedTest $speedTask.Result
        $speedTask = $null
    }
    $completed = $true
} finally {
    if ($speedTask) { [SpeedTestV1]::Cancel() }
    $writer.Dispose()
    if ($speedWriter) { $speedWriter.Dispose() }

    # ---------- Sammanfattning (kors aven vid Ctrl+C) ----------
    if ($tick -gt 0) {
        $state = if ($completed) { '' } else { ' - AVBRUTET' }
        Write-Host ("`n=== Sammanfattning: {0} matningar pa {1} s, {2} kort{3} ===" -f $tick,
            [math]::Round($clock.Elapsed.TotalSeconds, 0), $nics.Count, $state) -ForegroundColor Cyan
        Write-Host "Logg sparad: $LogPath"
        if ($speedOn) { Write-Host "Hastighet:   $SpeedLogPath" }
        if ($overruns) {
            Write-Host "Overskridna intervall: $overruns (matningen tog langre an $IntervalMs ms)" -ForegroundColor Yellow
        }

        Write-Host "`nLatens utan last (ms, median/p95/max):"
        $nics | ForEach-Object {
            $s = $_.Idle
            [pscustomobject]@{
                Kort         = $_.Name
                Typ          = $_.Typ
                Matningar    = $s.N
                Toppar       = $_.Spikes.Count
                'GW forl%'   = if ($s.N) { [math]::Round(100.0 * $s.GwLost / $s.N, 1) } else { '-' }
                'GW ms'      = Format-Stats $s.GwOk
                'Inet forl%' = if ($s.N) { [math]::Round(100.0 * $s.InLost / $s.N, 1) } else { '-' }
                'Inet ms'    = Format-Stats $s.InOk
            }
        } | Format-Table -AutoSize | Out-Host

        if ($speedOn) {
            Write-Host "Hastighet (Mbit/s, median / min-max over alla tester):"
            $nics | ForEach-Object {
                $ok = @($_.Speeds | Where-Object { -not $_.Error })
                $dn = @($_.Speeds | Where-Object { $_.DownMbps -ge 0 } | ForEach-Object { [math]::Round($_.DownMbps, 1) })
                $up = @($_.Speeds | Where-Object { $_.UpMbps -ge 0 } | ForEach-Object { [math]::Round($_.UpMbps, 1) })
                [pscustomobject]@{
                    Kort    = $_.Name
                    Tester  = $_.Speeds.Count
                    Fel     = $_.Speeds.Count - $ok.Count
                    Ned     = if ($dn.Count) { '{0} / {1}-{2}' -f (Get-Percentile $dn 0.5), ($dn | Measure-Object -Minimum).Minimum, ($dn | Measure-Object -Maximum).Maximum } else { '-' }
                    Upp     = if ($up.Count) { '{0} / {1}-{2}' -f (Get-Percentile $up 0.5), ($up | Measure-Object -Minimum).Minimum, ($up | Measure-Object -Maximum).Maximum } else { '-' }
                    'GW ms under last'   = Format-Stats $_.Load.GwOk
                    'Inet ms under last' = Format-Stats $_.Load.InOk
                }
            } | Format-Table -AutoSize | Out-Host
            Write-Host "Latens under last = medan hastighetstestet pagick (hog = bufferbloat i router/lank)."
        }

        foreach ($n in ($nics | Where-Object IsWlan)) {
            Write-Host "`n--- Wi-Fi: $($n.Name) ---" -ForegroundColor Cyan
            Write-Host "BSSID-byten (roaming): $($n.Roams.Count)"
            if ($n.Aps.Count) {
                $n.Aps.GetEnumerator() | ForEach-Object {
                    [pscustomobject]@{
                        BSSID       = $_.Key
                        Matningar   = $_.Value.N
                        SignalSnitt = if ($_.Value.SigN) { [math]::Round($_.Value.SigSum / $_.Value.SigN, 0) } else { '-' }
                        SignalMin   = if ($_.Value.SigN) { $_.Value.SigMin } else { '-' }
                    }
                } | Sort-Object Matningar -Descending | Format-Table -AutoSize | Out-Host
            }
            if ($n.Spikes.Count) {
                if ($n.SigSpikeN -and $n.SigNormN) {
                    Write-Host ("Signal i snitt vid toppar: {0}% | vid normala matningar: {1}%" -f
                        [math]::Round($n.SigSpikeSum / $n.SigSpikeN, 0), [math]::Round($n.SigNormSum / $n.SigNormN, 0))
                }
                $near = 0
                foreach ($s in $n.Spikes) {
                    foreach ($r in $n.Roams) {
                        if ([math]::Abs(($tickTimes[$s] - $tickTimes[$r]).TotalSeconds) -le $RoamWindowSec) { $near++; break }
                    }
                }
                Write-Host "Toppar inom +/-$RoamWindowSec s fran ett BSSID-byte: $near av $($n.Spikes.Count)"
            }
        }

        if ($nics.Count -gt 1 -and $spikeTicks.Count) {
            $all = @($spikeTicks.Values | Where-Object { $_ -eq $nics.Count }).Count
            Write-Host "`n--- Toppar jamfort mellan kort ---" -ForegroundColor Cyan
            Write-Host "Matningar med topp pa ALLA $($nics.Count) kort samtidigt: $all av $($spikeTicks.Count)"
            foreach ($n in $nics) {
                $only = @($n.Spikes | Where-Object { $spikeTicks[$_] -eq 1 }).Count
                Write-Host "Toppar bara pa '$($n.Name)': $only"
            }
        }

        $totalSpikes = @($nics | ForEach-Object { $_.Spikes.Count } | Measure-Object -Sum).Sum
        if ($totalSpikes) {
            Write-Host "`nTolkning (indikation, inte bevis):"
            Write-Host " - Topp samtidigt pa alla kort = gemensam orsak: router/gateway, internetlinjen eller datorn sjalv."
            Write-Host " - Topp bara pa ett kort = den lanken: for Wi-Fi roaming, lag signal eller kanalstorning."
            Write-Host " - Toppar nara BSSID-byte = roaming. Lagre signal vid toppar = rackvidd/storning."
            Write-Host " - Topp bara mot internet men inte gateway = bortom routern (ISP/uppstroms)."
        } else {
            Write-Host "`nInga toppar under testet (matningar under hastighetstest raknas inte)."
        }
    }
}
