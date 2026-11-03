#Requires -Version 5.1

<#
.SYNOPSIS
    Övervakar anslutningskvalitet (latens, jitter, paketförlust) mot en eller flera mål och loggar till CSV.
.DESCRIPTION
    Skickar en serie ICMP-pingar per mätintervall och räknar ut min/medel/max-latens, jitter och paketförlust.
    Default gateway läggs automatiskt till som mål så att man kan skilja på lokala problem (LAN/WiFi)
    och problem längre ut (ISP/internet).

    Filer som skapas i -LogDirectory (en uppsättning per körning, namngiven efter starttiden):
      ConnQuality_<dator>_<yyyyMMdd_HHmm>.csv  - en rad per mål och mätning
      ConnEvents_<dator>_<yyyyMMdd_HHmm>.log   - statusändringar (OK -> DOWN osv.) + start/stopp/sammanfattning

    Avbryt när som helst med Ctrl+C - en sammanfattning skrivs ändå.
.PARAMETER Target
    Ett eller flera mål (IP eller värdnamn). Default: 8.8.8.8 och 1.1.1.1.
.PARAMETER SkipGateway
    Lägg inte till default gateway som mål automatiskt.
.PARAMETER DurationHours
    Hur länge övervakningen ska köra. Default 24 timmar.
.PARAMETER IntervalSeconds
    Tid mellan mätningar (start till start). Default 10 sekunder.
.PARAMETER PingsPerSample
    Antal pingar per mål och mätning. Fler pingar = bättre jitter/förlust-statistik. Default 5.
.PARAMETER TimeoutMs
    Timeout per ping i millisekunder. Default 1000.
.PARAMETER HighLatencyMs
    Medellatens över detta värde markeras som DEGRADED. Default 100.
.PARAMETER LossWarnPercent
    Paketförlust (%) på eller över detta värde markeras som DEGRADED. Default 20 (= 1 av 5).
.PARAMETER LogDirectory
    Katalog för loggfiler. Skapas om den saknas. Default C:\Temp\ConnMonitor.
.PARAMETER Delimiter
    CSV-avgränsare. Default är systemets listavgränsare (';' på svenska Windows), vilket matchar
    decimaltecknet så att filen öppnas direkt i Excel.
.NOTES
    Version: 1.1.1
    Datum:   2026-11-03

    Ändringslogg:
      1.1.1  2026-11-03  Rättat inklistrad rad i hjälptexten (.DESCRIPTION).
      1.1.0  2026-11-03  Loggfiler namnges efter starttid (yyyyMMdd_HHmm) - egna filer per körning.
                         Versionsnummer skrivs i händelseloggen vid start.
      1.0.0  2026-11-03  Första version: latens, jitter, paketförlust, gateway-detektering,
                         händelselogg vid statusändring, sammanfattning, PS 5.1/7-kompatibel.

    Kom ihåg att uppdatera både .NOTES och $ScriptVersion vid ändringar.
.EXAMPLE
    .\Watch-ConnectionQuality.ps1
    Kör i 24 h mot gateway, 8.8.8.8 och 1.1.1.1.
.EXAMPLE
    .\Watch-ConnectionQuality.ps1 -Target 10.0.0.5,fileserver01 -DurationHours 8 -IntervalSeconds 5 -Delimiter ';'
.EXAMPLE
    .\Watch-ConnectionQuality.ps1 -DurationHours 0.05 -Verbose
    Snabbtest i 3 minuter.
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

    # Default = systemets listavgränsare (';' på svenska Windows) så att CSV:n öppnas rätt i Excel
    [string]$Delimiter = (Get-Culture).TextInfo.ListSeparator
)

# --- Variables ---
$ScriptVersion = '1.1.1'   # Håll i synk med .NOTES
$ComputerName = $env:COMPUTERNAME
# Starttid i filnamnet ger egna filer per körning (krockar inte heller om flera instanser körs parallellt)
$RunStamp     = Get-Date -Format 'yyyyMMdd_HHmm'
$CsvPath      = Join-Path $LogDirectory "ConnQuality_${ComputerName}_$RunStamp.csv"
$EventLogPath = Join-Path $LogDirectory "ConnEvents_${ComputerName}_$RunStamp.log"

# Statistik per mål för sammanfattningen (fylls på under körningen)
$Stats = @{}

# --- Functions ---
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Write-Verbose $entry
    Add-Content -Path $EventLogPath -Value $entry -Encoding UTF8
}

function Get-DefaultGateway {
    # Hämtar IPv4-gateway från den route som har lägst metric (= den som faktiskt används)
    try {
        $route = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop |
            Sort-Object { $_.RouteMetric + $_.InterfaceMetric } |
            Select-Object -First 1
        if ($route -and $route.NextHop -ne '0.0.0.0') { return $route.NextHop }
    }
    catch {
        Write-Log "Kunde inte hämta default gateway: $_" -Level WARN
    }
    return $null
}

function Measure-Target {
    <#
        Skickar $Count pingar mot $ComputerName och returnerar ett objekt med statistik.
        Använder System.Net.NetworkInformation.Ping direkt i stället för Test-Connection,
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
                # Typiskt DNS-fel eller att nätverkskortet saknar anslutning
                $errorText = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $_.Exception.Message }
            }
            # Kort paus mellan pingar så att de inte går i en klump
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

        # Jitter = medelvärdet av skillnaden mellan på varandra följande svarstider (enligt RFC 3550-principen)
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
    # Samlar löpande statistik per mål för sammanfattningen
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

    # Statusändring -> händelselogg
    if ($s.LastStatus -and $s.LastStatus -ne $Result.Status) {
        $level = if ($Result.Status -eq 'OK') { 'INFO' } else { 'WARN' }
        if ($Result.Status -eq 'DOWN') { $s.Outages++ }
        $since = if ($s.StatusSince) { " (föregående status varade $([int]((Get-Date) - $s.StatusSince).TotalSeconds) s)" } else { '' }
        Write-Log "$($Result.Target): $($s.LastStatus) -> $($Result.Status). Loss $($Result.LossPct)%, Avg $($Result.AvgMs) ms$since" -Level $level
        $s.StatusSince = Get-Date
    }
    elseif (-not $s.LastStatus) {
        $s.StatusSince = Get-Date
        if ($Result.Status -eq 'DOWN') { $s.Outages++ }
    }
    $s.LastStatus = $Result.Status
}

function Get-Summary {
    foreach ($t in $Stats.Keys) {
        $s = $Stats[$t]
        [PSCustomObject]@{
            Target          = $t
            Samples         = $s.Samples
            LossPct         = if ($s.Sent) { [math]::Round((($s.Sent - $s.Received) / $s.Sent) * 100, 2) } else { $null }
            AvgMs           = if ($s.Received) { [math]::Round($s.LatencySum / $s.Received, 1) } else { $null }
            MaxMs           = if ($s.Received) { $s.MaxMs } else { $null }
            DegradedSamples = $s.DegradedSamples
            DownSamples     = $s.DownSamples
            Outages         = $s.Outages
        }
    }
}

# --- Main ---
try {
    if (-not (Test-Path $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    # Bygg mållistan - gateway först så att den syns överst
    $targets = [System.Collections.Generic.List[string]]::new()
    if (-not $SkipGateway) {
        $gw = Get-DefaultGateway
        if ($gw) { $targets.Add($gw) } else { Write-Log 'Ingen default gateway hittades - fortsätter utan.' -Level WARN }
    }
    foreach ($t in $Target) { if (-not $targets.Contains($t)) { $targets.Add($t) } }

    foreach ($t in $targets) {
        $Stats[$t] = @{
            Samples = 0; Sent = 0; Received = 0; LatencySum = 0.0; MaxMs = 0
            DegradedSamples = 0; DownSamples = 0; Outages = 0; LastStatus = $null; StatusSince = $null
        }
    }

    # Kolumnbredd för målnamn i live-vyn
    $targetWidth = ($targets | Measure-Object -Property Length -Maximum).Maximum

    $startTime = Get-Date
    $endTime   = $startTime.AddHours($DurationHours)

    Write-Log "Watch-ConnectionQuality v$ScriptVersion. Start på $ComputerName. Mål: $($targets -join ', '). Intervall $IntervalSeconds s, $PingsPerSample pingar/mätning, slut $($endTime.ToString('yyyy-MM-dd HH:mm'))."
    Write-Host "Watch-ConnectionQuality v$ScriptVersion - övervakar $($targets -join ', ') till $($endTime.ToString('yyyy-MM-dd HH:mm')). Ctrl+C för att avbryta." -ForegroundColor Cyan
    Write-Host "Loggar till $CsvPath" -ForegroundColor Cyan

    while ((Get-Date) -lt $endTime) {
        $loopStart = Get-Date

        foreach ($t in $targets) {
            $result = Measure-Target -ComputerName $t -Count $PingsPerSample -Timeout $TimeoutMs `
                                     -HighLatency $HighLatencyMs -LossWarn $LossWarnPercent

            $result | Export-Csv -Path $CsvPath -Append -NoTypeInformation -Delimiter $Delimiter -Encoding UTF8 -ErrorAction Stop
            Update-Stats -Result $result

            # Live-vy i konsolen (Write-Host används medvetet här - det är en interaktiv statusrad, inte data)
            $color = switch ($result.Status) { 'OK' { 'Green' } 'DEGRADED' { 'Yellow' } default { 'Red' } }
            $line  = "{0}  {1,-$targetWidth} {2,-9} loss {3,5}%  avg {4,6} ms  max {5,6} ms  jitter {6,5} ms {7}" -f `
                $result.Timestamp, $result.Target, $result.Status, $result.LossPct,
                $result.AvgMs, $result.MaxMs, $result.JitterMs, $result.Error
            Write-Host $line -ForegroundColor $color
        }

        # Vänta resten av intervallet (mätningen i sig tar tid, så vi drar av den)
        $elapsed   = ((Get-Date) - $loopStart).TotalSeconds
        $remaining = $IntervalSeconds - $elapsed
        if ($remaining -gt 0) { Start-Sleep -Milliseconds ([int]($remaining * 1000)) }
    }
}
catch {
    Write-Log "ERROR: $_" -Level ERROR
    throw
}
finally {
    # Körs även vid Ctrl+C
    if ($Stats.Count -gt 0) {
        $summary = Get-Summary
        Write-Log "Stopp. Sammanfattning:`n$($summary | Format-Table -AutoSize | Out-String)"
        Write-Host "`nSammanfattning:" -ForegroundColor Cyan
        $summary | Format-Table -AutoSize | Out-String | Write-Host
    }
}
