#Requires -Version 5.1

<#
.SYNOPSIS
    Overvakar anslutningskvalitet (latens, jitter, paketforlust) mot en eller flera mal och loggar till CSV.
.DESCRIPTION
    Skickar en serie ICMP-pingar per matintervall och raknar ut min/medel/max-latens, jitter och paketforlust.
    Default gateway laggs automatiskt till som mal sa att man kan skilja pa lokala problem (LAN/WiFi)
    och problem langre ut (ISP/internet).

    Filer som skapas i -LogDirectory (en uppsattning per korning, namngiven efter starttiden):
      ConnQuality_<dator>_<yyyyMMdd_HHmm>.csv  - en rad per mal och matning
      ConnEvents_<dator>_<yyyyMMdd_HHmm>.log   - statusandringar (OK -> DOWN osv.) + start/stopp/sammanfattning
      ConnSpeed_<dator>_<yyyyMMdd_HHmm>.csv    - hastighetstester, en rad per kort och test

    HASTIGHETSTEST: 10 s efter start och sedan var -SpeedTestIntervalMin:e minut mats ned- och
    uppladdning PER NATVERKSKORT (alla aktiva kort med IPv4 och default gateway, -SpeedTestAdapter
    filtrerar) mot speed.cloudflare.com, eller mot Ookla med -SpeedTestProvider Ookla. TCP-anslutningarna binds till kortets IP, sa varje kort
    mats over sin egen lank; korten testas ett i taget. Testet kors i bakgrunden. Pingar som gors
    medan det pagar far UnderLoad = True i CSV:n, raknas inte in i status/handelser/statistiken
    och redovisas separat som latens under last (bufferbloat). Varje test forbrukar data
    (ungefar linjehastighet x SpeedTestSec per riktning och kort).

    Avbryt nar som helst med Ctrl+C - en sammanfattning skrivs anda.
.PARAMETER Target
    Ett eller flera mal (IP eller vardnamn). Default: 8.8.8.8 och 1.1.1.1.
.PARAMETER SkipGateway
    Lagg inte till default gateway som mal automatiskt.
.PARAMETER DurationHours
    Hur lange overvakningen ska kora. Default 24 timmar.
.PARAMETER IntervalSeconds
    Tid mellan matningar (start till start). Default 10 sekunder.
.PARAMETER PingsPerSample
    Antal pingar per mal och matning. Fler pingar = battre jitter/forlust-statistik. Default 5.
.PARAMETER TimeoutMs
    Timeout per ping i millisekunder. Default 1000.
.PARAMETER HighLatencyMs
    Medellatens over detta varde markeras som DEGRADED. Default 100.
.PARAMETER LossWarnPercent
    Paketforlust (%) pa eller over detta varde markeras som DEGRADED. Default 20 (= 1 av 5).
.PARAMETER LogDirectory
    Katalog for loggfiler. Skapas om den saknas. Default C:\Temp\ConnMonitor.
.PARAMETER Delimiter
    CSV-avgransare. Default ar systemets listavgransare (';' pa svenska Windows), vilket matchar
    decimaltecknet sa att filen oppnas direkt i Excel.
.PARAMETER SpeedTestIntervalMin
    Minuter mellan hastighetstester. Forsta testet kors 10 s efter start. Default 180
    (var tredje timme). 0 = inga hastighetstester.
.PARAMETER SpeedTestAdapter
    Vilka natverkskort som hastighetstestas, wildcards tillatna. Default '*' = alla aktiva kort
    med IPv4-adress och default gateway. Exempel: -SpeedTestAdapter 'Wi-Fi*','Ethernet'
.PARAMETER SpeedTestSec
    Mattid per riktning (ned/upp) och kort, i sekunder. Default 8.
.PARAMETER SpeedTestStreams
    Antal parallella TCP-anslutningar per matning. Default 4.
.PARAMETER SpeedTestHost
    HTTPS-server med Cloudflare-kompatibla andpunkter /__down?bytes=N och /__up.
    Default speed.cloudflare.com. Cloudflare svarar 429 Too Many Requests om man kor manga
    tester tatt (sett efter ca 20 tester inom 1,5 h); det loggas som fel for det testet.
    Anvands bara med -SpeedTestProvider Cloudflare.
.PARAMETER SpeedTestProvider
    Vilken tjanst hastighetstestet kor mot:
      Cloudflare (default) - inbyggd matning mot -SpeedTestHost, inget att installera.
      Ookla                - officiella Speedtest CLI (speedtest.exe), bunden till kortets IP med
                             --ip. Installeras med: winget install Ookla.Speedtest.CLI
    Ookla valjer narmaste server sjalv (eller -OoklaServerId) och tar ca 25-40 s per kort.
    -SpeedTestSec och -SpeedTestStreams paverkar inte Ookla.
.PARAMETER OoklaPath
    Sokvag till Ooklas speedtest.exe. Utan den letas den i PATH, i WinGet-paketkatalogen och
    bredvid skriptet.
.PARAMETER OoklaServerId
    Ookla-serverns id (lista: speedtest.exe -L). 0 (default) = Ookla valjer.
.NOTES
    Version: 1.3.1
    Datum:   2026-11-19

    Andringslogg:
      1.3.1  2026-11-19  Mat- och hastighetsmotorn (C#) och kortvalet ligger nu i
                         NetProbe.Engine.ps1 i samma mapp (delas med Test-WifiConnection och
                         Invoke-NicSpeedTest) i stallet for som kopia i skriptet - kopiera
                         hela Network-mappen. Ingen funktionell andring.
      1.3.0  2026-11-10  Flyttad fran repots rot till Network/.
                         -SpeedTestProvider Cloudflare|Ookla (+ -OoklaPath, -OoklaServerId).
                         Sokvagarna till de sparade filerna skrivs ut nar korningen stoppas
                         (aven vid Ctrl+C). Last CSV (t.ex. oppen i Excel) stoppar inte langre
                         korningen i fortid - raderna vantar i minnet och skrivs nar filen ar
                         ledig (vid stopp till en _ej-sparat_-fil bredvid).
      1.2.0  2026-11-04  Hastighetstest ned/upp per natverkskort (-SpeedTestIntervalMin m.fl.),
                         egen CSV, latens under last redovisas separat. Samma matmotor som
                         Network/Test-WifiConnection.ps1.
      1.1.1  2026-11-03  Rattat inklistrad rad i hjalptexten (.DESCRIPTION).
      1.1.0  2026-11-03  Loggfiler namnges efter starttid (yyyyMMdd_HHmm) - egna filer per korning.
                         Versionsnummer skrivs i handelseloggen vid start.
      1.0.0  2026-11-03  Forsta version: latens, jitter, paketforlust, gateway-detektering,
                         handelselogg vid statusandring, sammanfattning, PS 5.1/7-kompatibel.

    Kom ihag att uppdatera bade .NOTES och $ScriptVersion vid andringar.
.EXAMPLE
    .\Watch-ConnectionQuality.ps1
    Kor i 24 h mot gateway, 8.8.8.8 och 1.1.1.1.
.EXAMPLE
    .\Watch-ConnectionQuality.ps1 -Target 10.0.0.5,fileserver01 -DurationHours 8 -IntervalSeconds 5 -Delimiter ';'
.EXAMPLE
    .\Watch-ConnectionQuality.ps1 -DurationHours 0.05 -Verbose
    Snabbtest i 3 minuter.
.EXAMPLE
    .\Watch-ConnectionQuality.ps1 -DurationHours 8 -SpeedTestIntervalMin 60 -SpeedTestAdapter 'Wi-Fi*'
    8 timmar, hastighetstest av Wi-Fi-korten varje timme.
.EXAMPLE
    .\Watch-ConnectionQuality.ps1 -SpeedTestIntervalMin 0
    Bara latens/jitter/forlust, inga hastighetstester (som v1.1).
.EXAMPLE
    .\Watch-ConnectionQuality.ps1 -DurationHours 2 -SpeedTestIntervalMin 30 -SpeedTestProvider Ookla
    Hastighetstest med Ooklas Speedtest CLI i stallet for Cloudflare.
#>

[CmdletBinding()]
param(
    [string[]]$Target = @('8.8.8.8', '1.1.1.1'),

    [switch]$SkipGateway,

    [ValidateRange(0.001, 720)]
    [double]$DurationHours = 24,

    [ValidateRange(1, 3600)]
    [int]$IntervalSeconds = 10,

    [ValidateRange(1, 100)]
    [int]$PingsPerSample = 5,

    [ValidateRange(100, 10000)]
    [int]$TimeoutMs = 1000,

    [int]$HighLatencyMs = 100,

    [ValidateRange(0, 100)]
    [double]$LossWarnPercent = 20,

    [string]$LogDirectory = 'C:\Temp\ConnMonitor',

    # Default = systemets listavgransare (';' pa svenska Windows) sa att CSV:n oppnas ratt i Excel
    [string]$Delimiter = (Get-Culture).TextInfo.ListSeparator,

    [ValidateRange(0, 10080)]
    [int]$SpeedTestIntervalMin = 180,

    [string[]]$SpeedTestAdapter = @('*'),

    [ValidateRange(3, 60)]
    [int]$SpeedTestSec = 8,

    [ValidateRange(1, 16)]
    [int]$SpeedTestStreams = 4,

    [string]$SpeedTestHost = 'speed.cloudflare.com',

    [ValidateSet('Cloudflare', 'Ookla')]
    [string]$SpeedTestProvider = 'Cloudflare',

    [string]$OoklaPath,

    [ValidateRange(0, 2147483647)]
    [int]$OoklaServerId = 0
)

# --- Variables ---
$ScriptVersion = '1.3.1'   # Hall i synk med .NOTES
$ComputerName = $env:COMPUTERNAME
# Starttid i filnamnet ger egna filer per korning (krockar inte heller om flera instanser kors parallellt)
$RunStamp     = Get-Date -Format 'yyyyMMdd_HHmm'
$CsvPath      = Join-Path $LogDirectory "ConnQuality_${ComputerName}_$RunStamp.csv"
$EventLogPath = Join-Path $LogDirectory "ConnEvents_${ComputerName}_$RunStamp.log"
$SpeedCsvPath = Join-Path $LogDirectory "ConnSpeed_${ComputerName}_$RunStamp.csv"

# Statistik per mal for sammanfattningen (fylls pa under korningen)
$Stats = @{}

# Hastighetstest
$SpeedNics    = @()                                           # kort som testas
$SpeedResults = [System.Collections.Generic.List[object]]::new()
$speedTask    = $null

# CSV-rader som vantar pa att en last fil (t.ex. oppen i Excel) blir ledig
$PendingRows = @{}
$LockedFiles = @{}
$ExtraFiles  = @()

# Ookla: ett kort i taget, en speedtest.exe-process at gangen
$OoklaTimeoutSec = 120
$Ookla = @{ Exe = $null; Active = $false; Queue = $null; Proc = $null; Idx = -1; Out = $null; Err = $null
            Started = $null; Results = $null }

# --- Delad motor: ICMP, hastighetstest per kort, kortval (NetProbe.Engine.ps1) ---
$engine = Join-Path $PSScriptRoot 'NetProbe.Engine.ps1'
if (-not (Test-Path -LiteralPath $engine)) { throw "NetProbe.Engine.ps1 saknas bredvid Watch-ConnectionQuality.ps1 - kopiera hela Network-mappen." }
. $engine

# --- Functions ---
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Write-Verbose $entry
    # Ett lasfel i loggfilen far aldrig stoppa overvakningen
    try { Add-Content -Path $EventLogPath -Value $entry -Encoding UTF8 -ErrorAction Stop }
    catch { Write-Warning "Kunde inte skriva till $EventLogPath ($($_.Exception.Message)): $entry" }
}

function Write-CsvRow {
    <#
        Lagger till en rad i en CSV-fil. Ar filen last (typiskt: oppen i Excel) sparas raden i
        minnet och skrivs nar filen ar ledig igen - overvakningen fortsatter. Tidigare (v1.2.0)
        stoppade ett lasfel hela korningen i fortid.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Row
    )

    if (-not $PendingRows.ContainsKey($Path)) { $PendingRows[$Path] = [System.Collections.Generic.List[object]]::new() }
    $q = $PendingRows[$Path]
    $q.Add($Row)
    try {
        $q.ToArray() | Export-Csv -Path $Path -Append -NoTypeInformation -Delimiter $Delimiter -Encoding UTF8 -ErrorAction Stop
        if ($LockedFiles.ContainsKey($Path)) {
            Write-Log "Kan skriva till $Path igen - $($q.Count) vantande rader skrivna."
            Write-Host "Kan skriva till $Path igen ($($q.Count) rader ikapp)." -ForegroundColor Cyan
            $LockedFiles.Remove($Path)
        }
        $q.Clear()
    }
    catch {
        if (-not $LockedFiles.ContainsKey($Path)) {
            $LockedFiles[$Path] = $true
            $msg = "Kan inte skriva till $Path ($($_.Exception.Message)). Ar filen oppen i Excel? Raderna sparas i minnet och skrivs nar filen ar ledig - stang filen eller oppna en kopia."
            Write-Log $msg -Level WARN
            Write-Host $msg -ForegroundColor Yellow
        }
    }
}

function Save-PendingRows {
    # Vid stopp: forsok en sista gang; ar filen fortfarande last skrivs raderna till en ny fil bredvid
    foreach ($path in @($PendingRows.Keys)) {
        $q = $PendingRows[$path]
        if (-not $q.Count) { continue }
        try {
            $q.ToArray() | Export-Csv -Path $path -Append -NoTypeInformation -Delimiter $Delimiter -Encoding UTF8 -ErrorAction Stop
        }
        catch {
            $alt = [System.IO.Path]::ChangeExtension($path, $null).TrimEnd('.') + "_ej-sparat_$(Get-Date -Format 'HHmmss').csv"
            try {
                $q.ToArray() | Export-Csv -Path $alt -NoTypeInformation -Delimiter $Delimiter -Encoding UTF8 -ErrorAction Stop
                Write-Log "$path var last vid stopp - $($q.Count) rader skrevs till $alt i stallet." -Level WARN
                $script:ExtraFiles += $alt
            }
            catch { Write-Log "Kunde inte spara $($q.Count) rader for $path : $_" -Level ERROR }
        }
        $q.Clear()
    }
}

function Get-DefaultGateway {
    # Hamtar IPv4-gateway fran den route som har lagst metric (= den som faktiskt anvands)
    try {
        $route = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop |
            Sort-Object { $_.RouteMetric + $_.InterfaceMetric } |
            Select-Object -First 1
        if ($route -and $route.NextHop -ne '0.0.0.0') { return $route.NextHop }
    }
    catch {
        Write-Log "Kunde inte hamta default gateway: $_" -Level WARN
    }
    return $null
}

function Measure-Target {
    <#
        Skickar $Count pingar mot $ComputerName och returnerar ett objekt med statistik.
        Anvander System.Net.NetworkInformation.Ping direkt i stallet for Test-Connection,
        eftersom Test-Connection beter sig olika i PS 5.1 och PS 7 (egenskapsnamn,
        och PS 7 kastar inte fel vid timeout).
    #>
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [int]$Count = 5,
        [int]$Timeout = 1000,
        [int]$HighLatency = 100,
        [double]$LossWarn = 20
    )

    $pinger = [System.Net.NetworkInformation.Ping]::new()
    $rtts   = [System.Collections.Generic.List[double]]::new()
    $errorText = $null

    try {
        for ($i = 0; $i -lt $Count; $i++) {
            try {
                $reply = $pinger.Send($ComputerName, $Timeout)
                if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                    $rtts.Add([double]$reply.RoundtripTime)
                }
                else {
                    $errorText = $reply.Status.ToString()
                }
            }
            catch [System.Net.NetworkInformation.PingException] {
                # Typiskt DNS-fel eller att natverkskortet saknar anslutning
                $errorText = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $_.Exception.Message }
            }
            # Kort paus mellan pingar sa att de inte gar i en klump
            if ($i -lt $Count - 1) { Start-Sleep -Milliseconds 200 }
        }
    }
    finally {
        $pinger.Dispose()
    }

    $received = $rtts.Count
    $lossPct  = [math]::Round((($Count - $received) / $Count) * 100, 1)

    $min = $avg = $max = $jitter = $null
    if ($received -gt 0) {
        $m   = $rtts | Measure-Object -Minimum -Maximum -Average
        $min = $m.Minimum
        $max = $m.Maximum
        $avg = [math]::Round($m.Average, 1)

        # Jitter = medelvardet av skillnaden mellan pa varandra foljande svarstider (enligt RFC 3550-principen)
        if ($received -gt 1) {
            $diffs = for ($j = 1; $j -lt $received; $j++) { [math]::Abs($rtts[$j] - $rtts[$j - 1]) }
            $jitter = [math]::Round(($diffs | Measure-Object -Average).Average, 1)
        }
        else {
            $jitter = 0
        }
    }

    $status = if ($received -eq 0) { 'DOWN' }
              elseif ($lossPct -ge $LossWarn -or $avg -gt $HighLatency) { 'DEGRADED' }
              else { 'OK' }

    [PSCustomObject]@{
        Timestamp  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Computer   = $env:COMPUTERNAME
        Target     = $ComputerName
        Status     = $status
        Sent       = $Count
        Received   = $received
        LossPct    = $lossPct
        MinMs      = $min
        AvgMs      = $avg
        MaxMs      = $max
        JitterMs   = $jitter
        Error      = if ($status -ne 'OK') { $errorText } else { $null }
    }
}

function Update-Stats {
    # Samlar lopande statistik per mal for sammanfattningen
    param([Parameter(Mandatory)]$Result)

    $s = $Stats[$Result.Target]
    $s.Samples++
    $s.Sent     += $Result.Sent
    $s.Received += $Result.Received
    if ($null -ne $Result.AvgMs) {
        $s.LatencySum += $Result.AvgMs * $Result.Received
        if ($Result.MaxMs -gt $s.MaxMs) { $s.MaxMs = $Result.MaxMs }
    }
    if ($Result.Status -eq 'DOWN')     { $s.DownSamples++ }
    if ($Result.Status -eq 'DEGRADED') { $s.DegradedSamples++ }

    # Statusandring -> handelselogg
    if ($s.LastStatus -and $s.LastStatus -ne $Result.Status) {
        $level = if ($Result.Status -eq 'OK') { 'INFO' } else { 'WARN' }
        if ($Result.Status -eq 'DOWN') { $s.Outages++ }
        $since = if ($s.StatusSince) { " (foregaende status varade $([int]((Get-Date) - $s.StatusSince).TotalSeconds) s)" } else { '' }
        Write-Log "$($Result.Target): $($s.LastStatus) -> $($Result.Status). Loss $($Result.LossPct)%, Avg $($Result.AvgMs) ms$since" -Level $level
        $s.StatusSince = Get-Date
    }
    elseif (-not $s.LastStatus) {
        $s.StatusSince = Get-Date
        if ($Result.Status -eq 'DOWN') { $s.Outages++ }
    }
    $s.LastStatus = $Result.Status
}

function Update-LoadStats {
    # Pingar under hastighetstest: egen statistik, paverkar inte status/handelser
    param([Parameter(Mandatory)]$Result)

    $s = $Stats[$Result.Target]
    $s.LoadSamples++
    $s.LoadSent     += $Result.Sent
    $s.LoadReceived += $Result.Received
    if ($null -ne $Result.AvgMs) {
        $s.LoadLatencySum += $Result.AvgMs * $Result.Received
        if ($Result.MaxMs -gt $s.LoadMaxMs) { $s.LoadMaxMs = $Result.MaxMs }
    }
}

function Get-Summary {
    foreach ($t in $Stats.Keys) {
        $s = $Stats[$t]
        $row = [ordered]@{
            Target          = $t
            Samples         = $s.Samples
            LossPct         = if ($s.Sent) { [math]::Round((($s.Sent - $s.Received) / $s.Sent) * 100, 2) } else { $null }
            AvgMs           = if ($s.Received) { [math]::Round($s.LatencySum / $s.Received, 1) } else { $null }
            MaxMs           = if ($s.Received) { $s.MaxMs } else { $null }
            DegradedSamples = $s.DegradedSamples
            DownSamples     = $s.DownSamples
            Outages         = $s.Outages
        }
        if ($SpeedNics.Count) {
            $row.AvgMsUnderLoad = if ($s.LoadReceived) { [math]::Round($s.LoadLatencySum / $s.LoadReceived, 1) } else { $null }
            $row.MaxMsUnderLoad = if ($s.LoadReceived) { $s.LoadMaxMs } else { $null }
        }
        [PSCustomObject]$row
    }
}

function Get-SpeedTestNic {
    # Aktiva kort med anvandbar IPv4-adress och egen default gateway (Get-NetProbeNic i NetProbe.Engine.ps1)
    param([string[]]$Name = @('*'))

    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($c in @(Get-NetProbeNic -Name $Name)) {
        if ($c.SkipReason) { Write-Log "Hastighetstest: hoppar over '$($c.Name)' ($($c.SkipReason))." -Level WARN; continue }
        $list.Add([PSCustomObject]@{ Name = $c.Name; Src = $c.Src; LinkSpeed = $c.LinkSpeed })
    }
    return , $list.ToArray()
}

function Complete-SpeedTest {
    # Skriver resultaten fran ett avslutat hastighetstest till CSV, handelselogg och konsol
    param([Parameter(Mandatory)]$Results)

    foreach ($r in $Results) {
        $nic = $SpeedNics[$r.Index]
        # Ookla-resultat (PSCustomObject) har egen Server/ResultUrl; Cloudflare-resultat (C#) har det inte
        $server = if ($r.PSObject.Properties['Server'] -and $r.Server) { $r.Server } else { $SpeedTestHost }
        $url    = if ($r.PSObject.Properties['ResultUrl'] -and $r.ResultUrl) { $r.ResultUrl } else { $null }
        $row = [PSCustomObject]@{
            Timestamp = $r.Started.ToString('yyyy-MM-dd HH:mm:ss')
            Computer  = $ComputerName
            Adapter   = $nic.Name
            SourceIP  = $nic.Src
            LinkSpeed = $nic.LinkSpeed
            Provider  = $SpeedTestProvider
            Server    = $server
            DownMbps  = if ($r.DownMbps -ge 0) { [math]::Round($r.DownMbps, 1) } else { $null }
            UpMbps    = if ($r.UpMbps -ge 0) { [math]::Round($r.UpMbps, 1) } else { $null }
            DownMB    = [math]::Round($r.DownBytes / 1MB, 1)
            UpMB      = [math]::Round($r.UpBytes / 1MB, 1)
            Status    = if ($r.Error) { 'ERROR' } else { 'OK' }
            Error     = if ($r.Error) { $r.Error } else { $null }
            ResultUrl = $url
        }
        Write-CsvRow -Path $SpeedCsvPath -Row $row
        $SpeedResults.Add($row)

        $down = if ($null -ne $row.DownMbps) { $row.DownMbps } else { 'fel' }
        $up   = if ($null -ne $row.UpMbps) { $row.UpMbps } else { 'fel' }
        $etxt = if ($r.Error) { " ($($r.Error))" } else { '' }
        $msg  = "Hastighet [$($nic.Name)] fran $($nic.Src): ned $down Mbit/s, upp $up Mbit/s$etxt"
        Write-Log $msg -Level $(if ($r.Error) { 'WARN' } else { 'INFO' })
        Write-Host "$($row.Timestamp)  $msg" -ForegroundColor $(if ($r.Error) { 'Yellow' } else { 'Green' })
    }
}

function Find-OoklaCli {
    # Letar upp Ooklas speedtest.exe och kontrollerar att det ar Ooklas (inte t.ex. Python-klonen speedtest-cli)
    param([string]$Path)

    $cands = [System.Collections.Generic.List[string]]::new()
    if ($Path) { $cands.Add($Path) }
    else {
        foreach ($c in @(Get-Command speedtest.exe -All -CommandType Application -ErrorAction SilentlyContinue)) { $cands.Add($c.Source) }
        # Ingen -Recurse: wildcard + -Recurse i PS 5.1 soker igenom hela disken
        foreach ($f in @(Get-ChildItem -Path "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\Ookla.Speedtest.CLI*\speedtest.exe" -ErrorAction SilentlyContinue)) { $cands.Add($f.FullName) }
        if ($PSScriptRoot) { $cands.Add((Join-Path $PSScriptRoot 'speedtest.exe')) }
    }
    foreach ($p in $cands) {
        if (-not $p -or -not (Test-Path -LiteralPath $p)) { continue }
        try {
            $v = (& $p --version 2>&1 | Out-String)
            if ($v -match 'Ookla') { return $p }
            Write-Log "Hastighetstest: '$p' ar inte Ooklas Speedtest CLI - hoppar over." -Level WARN
        }
        catch {
            Write-Log "Hastighetstest: kunde inte kora '$p': $_" -Level WARN
        }
    }
    return $null
}

function Start-OoklaNic {
    # Startar speedtest.exe for nasta kort i kon (kors i bakgrunden, utdata till temp-filer)
    $Ookla.Idx     = $Ookla.Queue.Dequeue()
    $src           = $SpeedNics[$Ookla.Idx].Src
    $base          = Join-Path $env:TEMP "WatchConnQuality_ookla_${PID}_$($Ookla.Idx)"
    $Ookla.Out     = "$base.out"
    $Ookla.Err     = "$base.err"
    $Ookla.Started = Get-Date
    $argList = @('--accept-license', '--accept-gdpr', '--format=json', '--progress=no', "--ip=$src")
    if ($OoklaServerId -gt 0) { $argList += "--server-id=$OoklaServerId" }
    try {
        $Ookla.Proc = Start-Process -FilePath $Ookla.Exe -ArgumentList $argList -NoNewWindow -PassThru `
                          -RedirectStandardOutput $Ookla.Out -RedirectStandardError $Ookla.Err -ErrorAction Stop
        $null = $Ookla.Proc.Handle   # PS 5.1: behall handtaget sa att HasExited/ExitCode gar att lasa
    }
    catch {
        $Ookla.Proc = $null
        $Ookla.Results.Add((Read-OoklaResult -StartError "$_"))
    }
}

function Read-OoklaResult {
    # Tolkar speedtest.exe:s JSON-utdata (en rad per objekt; type=result ar slutresultatet)
    param([switch]$TimedOut, [string]$StartError)

    $r = [PSCustomObject]@{
        Index = $Ookla.Idx; Src = $SpeedNics[$Ookla.Idx].Src; Started = $Ookla.Started
        DownMbps = -1; UpMbps = -1; DownBytes = 0; UpBytes = 0; Error = ''; Server = ''; ResultUrl = ''
    }
    if ($StartError) { $r.Error = "kunde inte starta speedtest.exe: $StartError"; return $r }

    $res = $null; $lastErr = $null
    foreach ($line in @(Get-Content -LiteralPath $Ookla.Out -ErrorAction SilentlyContinue)) {
        if ($line -notmatch '^\s*\{') { continue }
        try { $o = $line | ConvertFrom-Json -ErrorAction Stop } catch { continue }
        if ($o.type -eq 'result') { $res = $o }
        elseif ($o.level -eq 'error' -and $o.message -notmatch '^bind\(') { $lastErr = $o.message }
        elseif ($o.error) { $lastErr = $o.error }
    }

    if ($res) {
        $r.DownMbps  = [double]$res.download.bandwidth * 8 / 1e6
        $r.UpMbps    = [double]$res.upload.bandwidth * 8 / 1e6
        $r.DownBytes = [long]$res.download.bytes
        $r.UpBytes   = [long]$res.upload.bytes
        $r.Server    = "$($res.server.name), $($res.server.location) (id $($res.server.id))"
        if ($res.result.url) { $r.ResultUrl = $res.result.url }
        # Kontroll att testet faktiskt gick over ratt kort
        $used = $res.interface.internalIp
        if ($used -and $used -ne $r.Src) { $r.Error = "Ookla anvande $used i stallet for $($r.Src)" }
    }
    else {
        $errText = (Get-Content -LiteralPath $Ookla.Err -ErrorAction SilentlyContinue | Where-Object { $_ }) -join ' '
        $r.Error = if ($TimedOut) { "timeout efter $OoklaTimeoutSec s" }
                   elseif ($lastErr) { $lastErr }
                   elseif ($errText) { $errText }
                   else { 'inget resultat fran speedtest.exe' }
    }
    foreach ($f in $Ookla.Out, $Ookla.Err) { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue }
    return $r
}

function Step-OoklaRun {
    # Anropas varje varv: skordar ett klart kort och startar nasta. -NoMore = starta inga fler kort.
    param([switch]$NoMore)

    if (-not $Ookla.Active) { return }
    if ($Ookla.Proc) {
        $timedOut = ((Get-Date) - $Ookla.Started).TotalSeconds -gt $OoklaTimeoutSec
        if (-not $Ookla.Proc.HasExited -and -not $timedOut) { return }
        if (-not $Ookla.Proc.HasExited) {
            try { $Ookla.Proc.Kill() } catch { Write-Verbose "Kill: $_" }
            $null = $Ookla.Proc.WaitForExit(5000)
        }
        $Ookla.Results.Add((Read-OoklaResult -TimedOut:$timedOut))
        $Ookla.Proc = $null
    }
    if ($NoMore -and $Ookla.Queue.Count) {
        Write-Log "Hastighetstest: tiden slut - $($Ookla.Queue.Count) kort testades inte i sista omgangen." -Level WARN
        $Ookla.Queue.Clear()
    }
    if ($Ookla.Queue.Count) { Start-OoklaNic; if ($Ookla.Proc) { return } ; Step-OoklaRun -NoMore:$NoMore; return }

    $Ookla.Active = $false
    if ($Ookla.Results.Count) { Complete-SpeedTest -Results $Ookla.Results.ToArray() }
}

function Get-Median([double[]]$Values) {
    if ($null -eq $Values -or $Values.Count -eq 0) { return $null }
    $a = @($Values | Sort-Object)
    $m = [int][math]::Floor($a.Count / 2)
    if ($a.Count % 2) { return $a[$m] }
    return [math]::Round(($a[$m - 1] + $a[$m]) / 2, 1)
}

function Get-SpeedSummary {
    foreach ($nic in $SpeedNics) {
        $rows = @($SpeedResults | Where-Object { $_.Adapter -eq $nic.Name })
        $dn = [double[]]@($rows | Where-Object { $null -ne $_.DownMbps } | ForEach-Object { $_.DownMbps })
        $up = [double[]]@($rows | Where-Object { $null -ne $_.UpMbps } | ForEach-Object { $_.UpMbps })
        [PSCustomObject]@{
            Adapter      = $nic.Name
            Tests        = $rows.Count
            Errors       = @($rows | Where-Object { $_.Status -ne 'OK' }).Count
            DownMedian   = Get-Median $dn
            DownMin      = if ($dn.Count) { ($dn | Measure-Object -Minimum).Minimum } else { $null }
            DownMax      = if ($dn.Count) { ($dn | Measure-Object -Maximum).Maximum } else { $null }
            UpMedian     = Get-Median $up
            UpMin        = if ($up.Count) { ($up | Measure-Object -Minimum).Minimum } else { $null }
            UpMax        = if ($up.Count) { ($up | Measure-Object -Maximum).Maximum } else { $null }
        }
    }
}

# --- Main ---
try {
    if (-not (Test-Path $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    # Bygg mallistan - gateway forst sa att den syns overst
    $targets = [System.Collections.Generic.List[string]]::new()
    if (-not $SkipGateway) {
        $gw = Get-DefaultGateway
        if ($gw) { $targets.Add($gw) } else { Write-Log 'Ingen default gateway hittades - fortsatter utan.' -Level WARN }
    }
    foreach ($t in $Target) { if (-not $targets.Contains($t)) { $targets.Add($t) } }

    foreach ($t in $targets) {
        $Stats[$t] = @{
            Samples = 0; Sent = 0; Received = 0; LatencySum = 0.0; MaxMs = 0
            DegradedSamples = 0; DownSamples = 0; Outages = 0; LastStatus = $null; StatusSince = $null
            LoadSamples = 0; LoadSent = 0; LoadReceived = 0; LoadLatencySum = 0.0; LoadMaxMs = 0
        }
    }

    # Kort att hastighetstesta
    if ($SpeedTestIntervalMin -gt 0) {
        $SpeedNics = Get-SpeedTestNic -Name $SpeedTestAdapter
        if (-not $SpeedNics.Count) {
            Write-Log "Hastighetstest: inga aktiva kort med IPv4 och default gateway matchar '$($SpeedTestAdapter -join "', '")' - hastighetstest av." -Level WARN
        }
        elseif ($SpeedTestProvider -eq 'Ookla') {
            $Ookla.Exe = Find-OoklaCli -Path $OoklaPath
            if (-not $Ookla.Exe) {
                $msg = 'Hastighetstest: hittar inte Ooklas speedtest.exe - hastighetstest av. Installera med "winget install Ookla.Speedtest.CLI" (ny konsol efterat) eller ange -OoklaPath.'
                Write-Log $msg -Level WARN
                Write-Host $msg -ForegroundColor Yellow
                $SpeedNics = @()
            }
        }
    }

    # Kolumnbredd for malnamn i live-vyn
    $targetWidth = ($targets | Measure-Object -Property Length -Maximum).Maximum

    $startTime   = Get-Date
    $endTime     = $startTime.AddHours($DurationHours)
    $nextSpeedAt = $startTime.AddSeconds(10)

    Write-Log "Watch-ConnectionQuality v$ScriptVersion. Start pa $ComputerName. Mal: $($targets -join ', '). Intervall $IntervalSeconds s, $PingsPerSample pingar/matning, slut $($endTime.ToString('yyyy-MM-dd HH:mm'))."
    Write-Host "Watch-ConnectionQuality v$ScriptVersion - overvakar $($targets -join ', ') till $($endTime.ToString('yyyy-MM-dd HH:mm')). Ctrl+C for att avbryta." -ForegroundColor Cyan
    Write-Host "Loggar till $CsvPath" -ForegroundColor Cyan
    if ($SpeedNics.Count) {
        $how = if ($SpeedTestProvider -eq 'Ookla') {
            $sid = if ($OoklaServerId -gt 0) { "server $OoklaServerId" } else { 'Ookla valjer server' }
            "Ookla ($($Ookla.Exe), $sid), ca 25-40 s per kort"
        }
        else { "$SpeedTestHost, $SpeedTestSec s ned + $SpeedTestSec s upp per kort, $SpeedTestStreams strommar" }
        $speedInfo = "Hastighetstest: $how, forst efter 10 s och sedan var $SpeedTestIntervalMin`:e minut. Kort: " +
            (($SpeedNics | ForEach-Object { "$($_.Name) ($($_.Src))" }) -join ', ')
        Write-Log $speedInfo
        Write-Host $speedInfo -ForegroundColor Cyan
        Write-Host "Hastighetstester till $SpeedCsvPath" -ForegroundColor Cyan
    }
    else {
        Write-Host 'Hastighetstest: av' -ForegroundColor Cyan
    }

    while ((Get-Date) -lt $endTime) {
        $loopStart = Get-Date

        # Hastighetstest: skorda ett avslutat, starta nasta nar det ar dags
        if ($SpeedNics.Count) {
            if ($SpeedTestProvider -eq 'Ookla') {
                Step-OoklaRun
                if (-not $Ookla.Active -and $loopStart -ge $nextSpeedAt) {
                    $Ookla.Queue   = [System.Collections.Generic.Queue[int]]::new([int[]](0..($SpeedNics.Count - 1)))
                    $Ookla.Results = [System.Collections.Generic.List[object]]::new()
                    $Ookla.Active  = $true
                    $msg = "Hastighetstest (Ookla) startat ($($SpeedNics.Count) kort, ca $($SpeedNics.Count * 30) s)"
                    Write-Log $msg
                    Write-Host "$($loopStart.ToString('yyyy-MM-dd HH:mm:ss'))  $msg" -ForegroundColor DarkCyan
                    Step-OoklaRun
                    while ($nextSpeedAt -le $loopStart) { $nextSpeedAt = $nextSpeedAt.AddMinutes($SpeedTestIntervalMin) }
                }
            }
            else {
                if ($speedTask -and $speedTask.IsCompleted) {
                    Complete-SpeedTest -Results $speedTask.Result
                    $speedTask = $null
                }
                if (-not $speedTask -and $loopStart -ge $nextSpeedAt) {
                    $speedTask = [SpeedTestV1]::StartAll([string[]]@($SpeedNics | ForEach-Object { $_.Src }),
                        $SpeedTestHost, $SpeedTestSec, $SpeedTestStreams)
                    $msg = "Hastighetstest startat ($($SpeedNics.Count) kort, ca $($SpeedNics.Count * 2 * $SpeedTestSec) s)"
                    Write-Log $msg
                    Write-Host "$($loopStart.ToString('yyyy-MM-dd HH:mm:ss'))  $msg" -ForegroundColor DarkCyan
                    while ($nextSpeedAt -le $loopStart) { $nextSpeedAt = $nextSpeedAt.AddMinutes($SpeedTestIntervalMin) }
                }
            }
        }
        $underLoad = ($null -ne $speedTask) -or $Ookla.Active

        foreach ($t in $targets) {
            $result = Measure-Target -ComputerName $t -Count $PingsPerSample -Timeout $TimeoutMs `
                                     -HighLatency $HighLatencyMs -LossWarn $LossWarnPercent
            $result | Add-Member -NotePropertyName UnderLoad -NotePropertyValue $underLoad

            Write-CsvRow -Path $CsvPath -Row $result
            # Hog latens under hastighetstest ar vantad - egen statistik, inga statusandringar
            if ($underLoad) { Update-LoadStats -Result $result } else { Update-Stats -Result $result }

            # Live-vy i konsolen (Write-Host anvands medvetet har - det ar en interaktiv statusrad, inte data)
            $color = switch ($result.Status) { 'OK' { 'Green' } 'DEGRADED' { 'Yellow' } default { 'Red' } }
            if ($underLoad) { $color = 'DarkGray' }
            $line  = "{0}  {1,-$targetWidth} {2,-9} loss {3,5}%  avg {4,6} ms  max {5,6} ms  jitter {6,5} ms {7}" -f `
                $result.Timestamp, $result.Target, $result.Status, $result.LossPct,
                $result.AvgMs, $result.MaxMs, $result.JitterMs, $result.Error
            if ($underLoad) { $line += ' [under hastighetstest]' }
            Write-Host $line -ForegroundColor $color
        }

        # Vanta resten av intervallet (matningen i sig tar tid, sa vi drar av den)
        $elapsed   = ((Get-Date) - $loopStart).TotalSeconds
        $remaining = $IntervalSeconds - $elapsed
        # ...men aldrig forbi sluttiden
        $remaining = [math]::Min($remaining, ($endTime - (Get-Date)).TotalSeconds)
        if ($remaining -gt 0) { Start-Sleep -Milliseconds ([int]($remaining * 1000)) }
    }

    # Ett pagaende hastighetstest far bli klart (det ar tidsbegransat)
    if ($speedTask) {
        Write-Host 'Vantar pa pagaende hastighetstest...' -ForegroundColor DarkCyan
        while (-not $speedTask.IsCompleted) { Start-Sleep -Milliseconds 200 }
        Complete-SpeedTest -Results $speedTask.Result
        $speedTask = $null
    }
    # Ookla: lat kortet som testas just nu bli klart (max $OoklaTimeoutSec s), starta inga fler
    if ($Ookla.Active) {
        Write-Host 'Vantar pa pagaende hastighetstest (Ookla)...' -ForegroundColor DarkCyan
        while ($Ookla.Active) { Step-OoklaRun -NoMore; if ($Ookla.Active) { Start-Sleep -Milliseconds 500 } }
    }
}
catch {
    Write-Log "ERROR: $_" -Level ERROR
    throw
}
finally {
    # Kors aven vid Ctrl+C
    if ($speedTask -and -not $speedTask.IsCompleted) {
        [SpeedTestV1]::Cancel()
        Write-Log 'Pagaende hastighetstest avbrutet.' -Level WARN
    }
    if ($Ookla.Proc -and -not $Ookla.Proc.HasExited) {
        try { $Ookla.Proc.Kill() } catch { Write-Verbose "Kill: $_" }
        foreach ($f in $Ookla.Out, $Ookla.Err) { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue }
        Write-Log 'Pagaende hastighetstest (Ookla) avbrutet.' -Level WARN
    }
    if ($Stats.Count -gt 0) {
        $summary = Get-Summary
        Write-Log "Stopp. Sammanfattning:`n$($summary | Format-Table -AutoSize | Out-String -Width 250)"
        Write-Host "`nSammanfattning:" -ForegroundColor Cyan
        $summary | Format-Table -AutoSize | Out-String | Write-Host
    }
    if ($SpeedNics.Count) {
        $speedSummary = Get-SpeedSummary
        Write-Log "Hastighet (Mbit/s):`n$($speedSummary | Format-Table -AutoSize | Out-String -Width 250)"
        Write-Host 'Hastighet (Mbit/s):' -ForegroundColor Cyan
        $speedSummary | Format-Table -AutoSize | Out-String | Write-Host
        if (@($Stats.Values | Where-Object { $_.LoadReceived }).Count) {
            Write-Host 'AvgMsUnderLoad/MaxMsUnderLoad = latens medan hastighetstestet pagick (hog = bufferbloat i router/lank).'
        }
    }
    # Var resultatet hamnade (aven vid Ctrl+C)
    Save-PendingRows
    $saved = @(@($CsvPath, $EventLogPath, $SpeedCsvPath) + $ExtraFiles | Where-Object { Test-Path -LiteralPath $_ })
    if ($saved.Count) {
        Write-Host "`nSparade filer:" -ForegroundColor Cyan
        foreach ($f in $saved) { Write-Host "  $f" -ForegroundColor Cyan }
    }
}
