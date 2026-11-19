<#
.SYNOPSIS
    Delat natverksmotor for Network-verktygen. Dot-sourcas, kors inte sjalvt.

.DESCRIPTION
    Innehaller det som tidigare fanns som exakta kopior i flera skript:

      [NetConnProbeV1]   Parallell ICMP echo med bunden kalladress (IcmpSendEcho2Ex i
                         iphlpapi.dll) - samma sak som 'ping -S', utan ping.exe och utan
                         sprakberoende texttolkning.
      [SpeedTestV1]      Hastighetstest ned/upp per kalladress mot en HTTPS-server med
                         Cloudflare-kompatibla andpunkter (/__down?bytes=N, /__up).
      Get-NetProbeNic    Aktiva kort med anvandbar IPv4-adress och egen default gateway.
      ConvertFrom-NetshWlan / Get-WlanSnapshot
                         Tolkar 'netsh wlan show interfaces' (svensk/engelsk Windows),
                         kopplat till kort via GUID.

    Anvands av:
      Network/Invoke-NicSpeedTest.ps1
      Network/Test-WifiConnection.ps1
      Network/Watch-ConnectionQuality.ps1

    Anvandning i ett skript i samma mapp:
      . (Join-Path $PSScriptRoot 'NetProbe.Engine.ps1')

    Typnamnen har versionssuffix (V1). Andras C#-koden pa ett satt som inte ar
    bakatkompatibelt maste suffixet hojas (V2), annars krockar det med en redan laddad
    typ i samma PowerShell-session (Add-Type kan inte ersatta en typ).

.NOTES
    Version : 1.0.0
    Changelog:
      1.0.0  2026-11-19  Bruten ur Test-WifiConnection 1.1.0 och Watch-ConnectionQuality
                         1.3.0, dar C#-blocket (297 rader) och kortvalet fanns som exakta
                         kopior. Ingen funktionell andring i C#-koden.
#>

# ---------- ICMP- och hastighetsmotor (C#) ----------
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

# ---------- Kort och Wi-Fi ----------
function ConvertFrom-NetshWlan {
    # Tolkar 'netsh wlan show interfaces'. Returnerar hashtable: GUID (gemener, utan {}) -> objekt.
    param([string[]]$Lines)
    $map = @{}
    $blocks = [System.Collections.Generic.List[hashtable]]::new()
    $cur = @{}
    foreach ($line in $Lines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            if ($cur.Count) { $blocks.Add($cur); $cur = @{} }
            continue
        }
        if ($line -match '^\s*(.+?)\s+:\s*(.*)$') {
            $k = $Matches[1].Trim()
            if (-not $cur.ContainsKey($k)) { $cur[$k] = $Matches[2].Trim() }
        }
    }
    if ($cur.Count) { $blocks.Add($cur) }

    foreach ($b in $blocks) {
        $pick = {
            param($pattern)
            foreach ($key in $b.Keys) { if ($key -match $pattern) { return $b[$key] } }
            return $null
        }
        $guid = & $pick '^GUID$'
        if (-not $guid) { continue }
        $bssid = & $pick '^BSSID$'
        if (-not $bssid) { $bssid = & $pick '^AP BSSID$' }
        $sig  = & $pick '^Signal$'
        $rssi = & $pick '^RSSI$'
        $map[$guid.Trim([char[]]'{}').ToLowerInvariant()] = [pscustomobject]@{
            Namn     = & $pick '^(Name|Namn)$'
            Status   = & $pick '^(State|Tillst.nd)$'
            SSID     = & $pick '^SSID$'
            BSSID    = if ($bssid) { $bssid.ToLowerInvariant() } else { $null }
            Signal   = if ($sig -match '(\d+)') { [int]$Matches[1] } else { $null }
            Rssi     = if ($rssi -match '(-?\d+)') { [int]$Matches[1] } else { $null }
            Kanal    = & $pick '^(Channel|Kanal)$'
            Band     = & $pick '^Band$'
            Radiotyp = & $pick '^(Radio type|Radiotyp)$'
            RxRate   = & $pick '^(Receive rate|Mottagningshastighet)'
            TxRate   = & $pick '^(Transmit rate|.verf.ringshastighet)'
        }
    }
    return $map
}

function Get-WlanSnapshot {
    # netsh kan skriva fel (t.ex. WLAN-tjansten stoppad) - fanga dem som text, kasta inte
    $ErrorActionPreference = 'Continue'
    $raw = @(& netsh.exe wlan show interfaces 2>&1 | ForEach-Object { "$_" })
    [pscustomobject]@{ Map = (ConvertFrom-NetshWlan -Lines $raw); Raw = $raw }
}

function Get-NetProbeNic {
    # Aktiva kort som matchar -Name (wildcards). Ett objekt per kort; SkipReason ar tom
    # for kort som gar att anvanda (IPv4-adress + egen default gateway), annars orsaken.
    [CmdletBinding()]
    param([string[]]$Name = @('*'))

    foreach ($a in @(Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' })) {
        $n = $a.Name
        if (-not @($Name | Where-Object { $n -like $_ }).Count) { continue }

        $isWlan = ($a.NdisPhysicalMedium -eq 9 -or $a.NdisPhysicalMedium -eq 1 -or "$($a.PhysicalMediaType)" -match '802\.11')
        $o = [ordered]@{
            Name       = $n
            Guid       = "$($a.InterfaceGuid)".Trim([char[]]'{}').ToLowerInvariant()
            IfIndex    = $a.ifIndex
            IsWlan     = $isWlan
            Typ        = if ($isWlan) { 'WLAN' } else { 'Kabel/ovrigt' }
            Src        = $null
            Gw         = $null
            LinkSpeed  = $a.LinkSpeed
            SkipReason = ''
        }

        $ip = Get-NetIPAddress -InterfaceIndex $a.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
              Where-Object { $_.IPAddress -notlike '169.254.*' -and "$($_.AddressState)" -eq 'Preferred' } |
              Select-Object -First 1
        if (-not $ip) {
            $o.SkipReason = 'ingen anvandbar IPv4-adress'
        } else {
            $o.Src = $ip.IPAddress
            # Kortets egen default route, lagst metric
            $route = Get-NetRoute -InterfaceIndex $a.ifIndex -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' `
                                  -PolicyStore ActiveStore -ErrorAction SilentlyContinue |
                     Where-Object { $_.NextHop -and $_.NextHop -ne '0.0.0.0' } |
                     Sort-Object RouteMetric | Select-Object -First 1
            if (-not $route) { $o.SkipReason = 'ingen default gateway pa kortet' } else { $o.Gw = $route.NextHop }
        }
        [PSCustomObject]$o
    }
}
