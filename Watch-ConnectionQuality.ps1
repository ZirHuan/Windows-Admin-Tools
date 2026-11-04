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
    filtrerar) mot speed.cloudflare.com. TCP-anslutningarna binds till kortets IP, sa varje kort
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
.NOTES
    Version: 1.2.0
    Datum:   2026-11-04

    Andringslogg:
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

    [string]$SpeedTestHost = 'speed.cloudflare.com'
)

# --- Variables ---
$ScriptVersion = '1.2.0'   # Hall i synk med .NOTES
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

# --- Hastighetstest per natverkskort (C#) ---
# EXAKT kopia av blocket i Network/Test-WifiConnection.ps1 (samma typnamn), sa att bada
# skripten kan koras i samma PowerShell-session. Andra i bada filerna samtidigt.
if (-not ('SpeedTestV1' -as [type])) {
    Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Security.Authentication;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

public static class NetConnProbeV1
{
    [DllImport("iphlpapi.dll", SetLastError = true)]
    static extern IntPtr IcmpCreateFile();

    [DllImport("iphlpapi.dll", SetLastError = true)]
    static extern bool IcmpCloseHandle(IntPtr handle);

    [DllImport("iphlpapi.dll", SetLastError = true)]
    static extern uint IcmpSendEcho2Ex(IntPtr handle, IntPtr evt, IntPtr apcRoutine, IntPtr apcContext,
        uint sourceAddress, uint destinationAddress, byte[] requestData, ushort requestSize,
        IntPtr requestOptions, IntPtr replyBuffer, uint replySize, uint timeout);

    static uint ToAddr(string ip)
    {
        return BitConverter.ToUInt32(IPAddress.Parse(ip).GetAddressBytes(), 0);
    }

    // >= 0 : round-trip time in ms
    //  < 0 : minus the IP status / Win32 error (11010 = timeout)
    public static int Probe(string src, string dst, int timeoutMs)
    {
        uint s, d;
        try { s = ToAddr(src); d = ToAddr(dst); }
        catch { return -87; }

        IntPtr h = IcmpCreateFile();
        if (h == IntPtr.Zero || h == new IntPtr(-1)) return -Marshal.GetLastWin32Error();

        byte[] data = new byte[32];
        for (int i = 0; i < data.Length; i++) data[i] = (byte)('a' + (i % 23));
        int replySize = 256 + data.Length;
        IntPtr buf = Marshal.AllocHGlobal(replySize);
        try
        {
            uint n = IcmpSendEcho2Ex(h, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, s, d,
                data, (ushort)data.Length, IntPtr.Zero, buf, (uint)replySize, (uint)timeoutMs);
            if (n == 0)
            {
                int err = Marshal.GetLastWin32Error();
                return err == 0 ? -11010 : -err;
            }
            // ICMP_ECHO_REPLY: Address @0, Status @4, RoundTripTime @8 (same on x86/x64)
            int status = Marshal.ReadInt32(buf, 4);
            if (status != 0) return -status;
            return Marshal.ReadInt32(buf, 8);
        }
        finally
        {
            Marshal.FreeHGlobal(buf);
            IcmpCloseHandle(h);
        }
    }

    // One dedicated thread per probe, so a slow probe never waits for a pool thread.
    public static Task<int>[] Start(string[] src, string[] dst, int timeoutMs)
    {
        Task<int>[] t = new Task<int>[src.Length];
        for (int i = 0; i < src.Length; i++)
        {
            string s = src[i];
            string d = dst[i];
            t[i] = Task.Factory.StartNew(() => Probe(s, d, timeoutMs), TaskCreationOptions.LongRunning);
        }
        return t;
    }

    public static int[] Collect(Task<int>[] t)
    {
        Task.WaitAll(t);
        int[] r = new int[t.Length];
        for (int i = 0; i < t.Length; i++) r[i] = t[i].Result;
        return r;
    }
}

public class SpeedResultV1
{
    public int Index;
    public string Src;
    public DateTime Started;
    public double DownMbps = -1;
    public double UpMbps = -1;
    public long DownBytes;
    public long UpBytes;
    public string Error = "";
}

public static class SpeedTestV1
{
    const long DownChunk = 50000000;   // bytes per GET (Cloudflare returns 403 at >= 100 MB)
    const long UpChunk = 25000000;     // bytes per POST
    static readonly byte[] Payload = MakePayload();
    static CancellationTokenSource cts;

    // Index of the NIC being tested right now, -1 when idle.
    public static volatile int Current = -1;

    static byte[] MakePayload()
    {
        byte[] b = new byte[65536];
        new Random(42).NextBytes(b);
        return b;
    }

    public static void Cancel()
    {
        CancellationTokenSource c = cts;
        if (c != null) c.Cancel();
    }

    // Tests each source address in turn (download, then upload).
    public static Task<SpeedResultV1[]> StartAll(string[] srcs, string host, int seconds, int streams)
    {
        cts = new CancellationTokenSource();
        CancellationToken tok = cts.Token;
        Current = 0;
        return Task.Factory.StartNew(() =>
        {
            SpeedResultV1[] res = new SpeedResultV1[srcs.Length];
            IPAddress ip = null;
            string dnsErr = null;
            try
            {
                foreach (IPAddress a in Dns.GetHostAddresses(host))
                    if (a.AddressFamily == AddressFamily.InterNetwork) { ip = a; break; }
                if (ip == null) dnsErr = "no IPv4 address for " + host;
            }
            catch (Exception e) { dnsErr = "DNS: " + e.Message; }

            for (int i = 0; i < srcs.Length; i++)
            {
                Current = i;
                SpeedResultV1 r = new SpeedResultV1();
                r.Index = i; r.Src = srcs[i]; r.Started = DateTime.Now;
                if (ip == null) { r.Error = dnsErr; }
                else if (tok.IsCancellationRequested) { r.Error = "cancelled"; }
                else
                {
                    string e1, e2;
                    long b1, b2;
                    r.DownMbps = Measure(true, srcs[i], ip, host, seconds, streams, tok, out b1, out e1);
                    r.UpMbps = Measure(false, srcs[i], ip, host, seconds, streams, tok, out b2, out e2);
                    r.DownBytes = b1; r.UpBytes = b2;
                    string err = "";
                    if (r.DownMbps < 0) err = "ned: " + e1;
                    if (r.UpMbps < 0) err = (err.Length > 0 ? err + " | " : "") + "upp: " + e2;
                    r.Error = err;
                }
                res[i] = r;
            }
            Current = -1;
            return res;
        }, TaskCreationOptions.LongRunning);
    }

    static double Measure(bool down, string src, IPAddress ip, string host, int seconds, int streams,
                          CancellationToken tok, out long totalBytes, out string error)
    {
        long[] counter = new long[1];
        string[] firstErr = new string[1];
        Stopwatch sw = Stopwatch.StartNew();
        long endMs = seconds * 1000L;
        long warmMs = Math.Min(1500L, endMs / 4);

        Thread[] th = new Thread[streams];
        for (int i = 0; i < streams; i++)
        {
            th[i] = new Thread(() =>
            {
                while (sw.ElapsedMilliseconds < endMs && !tok.IsCancellationRequested)
                {
                    try { Transfer(down, src, ip, host, sw, endMs, tok, counter); }
                    catch (Exception e)
                    {
                        if (firstErr[0] == null) firstErr[0] = e.GetBaseException().Message;
                        tok.WaitHandle.WaitOne(250);
                    }
                }
            });
            th[i].IsBackground = true;
            th[i].Start();
        }

        // Skip TCP slow start: measure from warmMs to endMs.
        tok.WaitHandle.WaitOne((int)warmMs);
        long b0 = Interlocked.Read(ref counter[0]);
        long t0 = sw.ElapsedMilliseconds;
        long rest = endMs - sw.ElapsedMilliseconds;
        if (rest > 0) tok.WaitHandle.WaitOne((int)rest);
        long b1 = Interlocked.Read(ref counter[0]);
        long t1 = sw.ElapsedMilliseconds;
        foreach (Thread t in th) t.Join(10000);

        totalBytes = Interlocked.Read(ref counter[0]);
        error = firstErr[0] ?? "";
        if (tok.IsCancellationRequested) { error = "cancelled"; return -1; }
        if (b1 - b0 <= 0 || t1 <= t0) { if (error.Length == 0) error = "no data"; return -1; }
        return (b1 - b0) * 8.0 / ((t1 - t0) / 1000.0) / 1000000.0;
    }

    static void Transfer(bool down, string src, IPAddress ip, string host, Stopwatch sw, long endMs,
                         CancellationToken tok, long[] counter)
    {
        using (Socket s = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp))
        {
            s.Bind(new IPEndPoint(IPAddress.Parse(src), 0));   // force the source NIC
            s.ReceiveTimeout = 5000;
            s.SendTimeout = 5000;
            s.NoDelay = true;
            if (!down) s.SendBufferSize = 262144;              // limits bytes counted but not yet sent
            s.Connect(new IPEndPoint(ip, 443));
            using (NetworkStream ns = new NetworkStream(s, false))
            using (SslStream ssl = new SslStream(ns, false))
            {
                ssl.AuthenticateAsClient(host, null, SslProtocols.Tls12, false);
                byte[] buf = new byte[65536];
                if (down)
                {
                    byte[] req = Encoding.ASCII.GetBytes("GET /__down?bytes=" + DownChunk + " HTTP/1.1\r\nHost: " + host +
                        "\r\nUser-Agent: Test-WifiConnection\r\nAccept: */*\r\nConnection: close\r\n\r\n");
                    ssl.Write(req, 0, req.Length);
                    MemoryStream head = new MemoryStream();
                    bool inHead = true;
                    while (sw.ElapsedMilliseconds < endMs && !tok.IsCancellationRequested)
                    {
                        int n = ssl.Read(buf, 0, buf.Length);
                        if (n <= 0) break;
                        if (inHead)
                        {
                            head.Write(buf, 0, n);
                            byte[] h = head.GetBuffer();
                            int len = (int)head.Length;
                            int idx = FindHeaderEnd(h, len);
                            if (idx < 0)
                            {
                                if (len > 32768) throw new IOException("HTTP header too large");
                                continue;
                            }
                            string status = Encoding.ASCII.GetString(h, 0, Math.Min(len, 64)).Split('\r')[0];
                            if (status.IndexOf(" 200") < 0) throw new IOException("HTTP " + status);
                            inHead = false;
                            int body = len - (idx + 4);
                            if (body > 0) Interlocked.Add(ref counter[0], body);
                            continue;
                        }
                        Interlocked.Add(ref counter[0], n);
                    }
                }
                else
                {
                    byte[] req = Encoding.ASCII.GetBytes("POST /__up HTTP/1.1\r\nHost: " + host +
                        "\r\nUser-Agent: Test-WifiConnection\r\nContent-Type: application/octet-stream\r\nContent-Length: " +
                        UpChunk + "\r\nConnection: close\r\n\r\n");
                    ssl.Write(req, 0, req.Length);
                    long sent = 0;
                    while (sent < UpChunk && sw.ElapsedMilliseconds < endMs && !tok.IsCancellationRequested)
                    {
                        int c = (int)Math.Min(Payload.Length, UpChunk - sent);
                        ssl.Write(Payload, 0, c);
                        sent += c;
                        Interlocked.Add(ref counter[0], c);
                    }
                    if (sent >= UpChunk)
                    {
                        // whole body sent: check the server accepted it
                        int n = ssl.Read(buf, 0, buf.Length);
                        if (n > 0)
                        {
                            string status = Encoding.ASCII.GetString(buf, 0, Math.Min(n, 64)).Split('\r')[0];
                            if (status.IndexOf(" 200") < 0) throw new IOException("HTTP " + status);
                        }
                    }
                }
            }
        }
    }

    static int FindHeaderEnd(byte[] b, int len)
    {
        for (int i = 0; i + 3 < len; i++)
            if (b[i] == 13 && b[i + 1] == 10 && b[i + 2] == 13 && b[i + 3] == 10) return i;
        return -1;
    }
}
'@
}

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
    # Aktiva kort med anvandbar IPv4-adress och egen default gateway (samma urval som Test-WifiConnection)
    param([string[]]$Name = @('*'))

    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($a in @(Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' })) {
        $n = $a.Name
        if (-not @($Name | Where-Object { $n -like $_ }).Count) { continue }

        $ip = Get-NetIPAddress -InterfaceIndex $a.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
              Where-Object { $_.IPAddress -notlike '169.254.*' -and "$($_.AddressState)" -eq 'Preferred' } |
              Select-Object -First 1
        if (-not $ip) { Write-Log "Hastighetstest: hoppar over '$n' (ingen anvandbar IPv4-adress)." -Level WARN; continue }

        $route = Get-NetRoute -InterfaceIndex $a.ifIndex -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' `
                              -PolicyStore ActiveStore -ErrorAction SilentlyContinue |
                 Where-Object { $_.NextHop -and $_.NextHop -ne '0.0.0.0' } | Select-Object -First 1
        if (-not $route) { Write-Log "Hastighetstest: hoppar over '$n' (ingen default gateway pa kortet)." -Level WARN; continue }

        $list.Add([PSCustomObject]@{ Name = $n; Src = $ip.IPAddress; LinkSpeed = $a.LinkSpeed })
    }
    return , $list.ToArray()
}

function Complete-SpeedTest {
    # Skriver resultaten fran ett avslutat hastighetstest till CSV, handelselogg och konsol
    param([Parameter(Mandatory)]$Results)

    foreach ($r in $Results) {
        $nic = $SpeedNics[$r.Index]
        $row = [PSCustomObject]@{
            Timestamp = $r.Started.ToString('yyyy-MM-dd HH:mm:ss')
            Computer  = $ComputerName
            Adapter   = $nic.Name
            SourceIP  = $nic.Src
            LinkSpeed = $nic.LinkSpeed
            Server    = $SpeedTestHost
            DownMbps  = if ($r.DownMbps -ge 0) { [math]::Round($r.DownMbps, 1) } else { $null }
            UpMbps    = if ($r.UpMbps -ge 0) { [math]::Round($r.UpMbps, 1) } else { $null }
            DownMB    = [math]::Round($r.DownBytes / 1MB, 1)
            UpMB      = [math]::Round($r.UpBytes / 1MB, 1)
            Status    = if ($r.Error) { 'ERROR' } else { 'OK' }
            Error     = if ($r.Error) { $r.Error } else { $null }
        }
        $row | Export-Csv -Path $SpeedCsvPath -Append -NoTypeInformation -Delimiter $Delimiter -Encoding UTF8 -ErrorAction Stop
        $SpeedResults.Add($row)

        $down = if ($null -ne $row.DownMbps) { $row.DownMbps } else { 'fel' }
        $up   = if ($null -ne $row.UpMbps) { $row.UpMbps } else { 'fel' }
        $etxt = if ($r.Error) { " ($($r.Error))" } else { '' }
        $msg  = "Hastighet [$($nic.Name)] fran $($nic.Src): ned $down Mbit/s, upp $up Mbit/s$etxt"
        Write-Log $msg -Level $(if ($r.Error) { 'WARN' } else { 'INFO' })
        Write-Host "$($row.Timestamp)  $msg" -ForegroundColor $(if ($r.Error) { 'Yellow' } else { 'Green' })
    }
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
        $speedInfo = "Hastighetstest: $SpeedTestHost, forst efter 10 s och sedan var $SpeedTestIntervalMin`:e minut, $SpeedTestSec s ned + $SpeedTestSec s upp per kort, $SpeedTestStreams strommar. Kort: " +
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
        $underLoad = $null -ne $speedTask

        foreach ($t in $targets) {
            $result = Measure-Target -ComputerName $t -Count $PingsPerSample -Timeout $TimeoutMs `
                                     -HighLatency $HighLatencyMs -LossWarn $LossWarnPercent
            $result | Add-Member -NotePropertyName UnderLoad -NotePropertyValue $underLoad

            $result | Export-Csv -Path $CsvPath -Append -NoTypeInformation -Delimiter $Delimiter -Encoding UTF8 -ErrorAction Stop
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
        if ($remaining -gt 0) { Start-Sleep -Milliseconds ([int]($remaining * 1000)) }
    }

    # Ett pagaende hastighetstest far bli klart (det ar tidsbegransat)
    if ($speedTask) {
        Write-Host 'Vantar pa pagaende hastighetstest...' -ForegroundColor DarkCyan
        while (-not $speedTask.IsCompleted) { Start-Sleep -Milliseconds 200 }
        Complete-SpeedTest -Results $speedTask.Result
        $speedTask = $null
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
}
