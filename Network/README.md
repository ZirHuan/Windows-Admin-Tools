# Network tools

Two PowerShell scripts for finding out **where** a bad connection goes wrong: on the
Wi-Fi link, on the LAN, in the router, or out at the ISP. Both run on Windows
PowerShell 5.1 and PowerShell 7, need no admin rights and no extra modules, and write
CSV files that open directly in Excel.

| Script | Use it when | Measures |
|---|---|---|
| [`Test-WifiConnection.ps1`](Test-WifiConnection.ps1) | You need a sharp, per-NIC picture: which adapter spikes, does it follow roaming or weak signal, what speed does each NIC really get | Latency to gateway + internet **per NIC** every second, Wi-Fi SSID/BSSID/signal/channel, speed test per NIC |
| [`Watch-ConnectionQuality.ps1`](Watch-ConnectionQuality.ps1) | You want a light, long-running monitor (hours to days) with an event log of outages | Latency, jitter and loss per target every 10 s, OK/DEGRADED/DOWN events, speed test per NIC |

Changes and test results: [CHANGELOG.md](CHANGELOG.md).

## How the speed test works (both scripts)

- Download and upload over HTTPS against `speed.cloudflare.com` (`-SpeedTestHost`
  takes any server with Cloudflare-compatible `/__down` and `/__up` endpoints).
- Every TCP stream is **bound to the NIC's own IPv4 address**, so each adapter is
  measured over its own link - not just whichever one Windows routes through.
- NICs are tested **one at a time** (they usually share the same internet line).
  Default: 4 parallel streams, 8 s per direction after a short TCP warm-up.
- Latency measured while a speed test runs is flagged and reported separately as
  *latency under load*. A large jump there means **bufferbloat** in the router or on
  the line - a common cause of lag in calls and games even when bandwidth is fine.
- Each test uses data: roughly line rate x `SpeedTestSec` per direction and NIC
  (about 0.5 GB per NIC on a 250 Mbit/s line with the defaults).
- Cloudflare answers `429 Too Many Requests` if you run many tests close together
  (seen after about 20 tests in 1.5 h). That direction is then logged as an error and
  the run carries on. Wait a while, or test less often.

## Test-WifiConnection.ps1

### Quick speed test only

```powershell
# All active NICs, once, then exit. No latency logging. ~16 s per NIC.
.\Test-WifiConnection.ps1 -SpeedTestOnly

# Only the Wi-Fi adapters, longer measurement for a steadier figure
.\Test-WifiConnection.ps1 -SpeedTestOnly -AdapterName 'Wi-Fi*' -SpeedTestSec 15
```

Example output:

```
=== Hastighet (Mbit/s) ===

Kort       Typ          Kalla         Lank    Ned   Upp   Signal Status
----       ---          -----         ----    ---   ---   ------ ------
Ethernet 3 Kabel/ovrigt 192.168.2.193 1 Gbps  231.3 221.8        ok
Ethernet   Kabel/ovrigt 192.168.2.141 10 Gbps 48.1  19           ok

Sparat: C:\Users\me\Desktop\netspeed-20261104-134537.csv
```

### Latency logging (with periodic speed tests)

```powershell
# 5 minutes, all NICs, one speed test per NIC after 10 s
.\Test-WifiConnection.ps1

# 24 hours, speed test every hour
.\Test-WifiConnection.ps1 -DurationSec 86400 -SpeedTestIntervalMin 60

# Only Wi-Fi for 15 minutes, spike threshold 25 ms, no speed tests
.\Test-WifiConnection.ps1 -AdapterName 'Wi-Fi' -DurationSec 900 -SpikeMs 25 -SpeedTestIntervalMin 0
```

Reading the summary:

- **Spike on every NIC at the same time** - common cause: router/gateway, the internet
  line, or the PC itself.
- **Spike on one NIC only** - that link. For Wi-Fi: roaming, weak signal or channel
  interference.
- **Spikes near a BSSID change** - roaming. Lower average signal at spikes - range or
  interference.
- **Internet spikes without gateway spikes** - beyond the router (ISP / upstream).

Files (Desktop by default, `-LogPath` to change): `netconn-<date>-<time>.csv` (one
row per NIC and measurement) and `netconn-<date>-<time>-hastighet.csv` (speed tests).
With `-SpeedTestOnly`: `netspeed-<date>-<time>.csv`.

## Watch-ConnectionQuality.ps1

```powershell
# 24 hours against the default gateway, 8.8.8.8 and 1.1.1.1, speed test every 3 h
.\Watch-ConnectionQuality.ps1

# Working day, speed test of the Wi-Fi adapter every hour
.\Watch-ConnectionQuality.ps1 -DurationHours 8 -SpeedTestIntervalMin 60 -SpeedTestAdapter 'Wi-Fi*'

# Own targets, latency only (the v1.1 behaviour)
.\Watch-ConnectionQuality.ps1 -Target 10.0.0.5,fileserver01 -SpeedTestIntervalMin 0

# 3-minute smoke test
.\Watch-ConnectionQuality.ps1 -DurationHours 0.05 -Verbose

# Speed test with Ookla's Speedtest CLI instead of Cloudflare (winget install Ookla.Speedtest.CLI)
.\Watch-ConnectionQuality.ps1 -DurationHours 2 -SpeedTestIntervalMin 30 -SpeedTestProvider Ookla
```

While a speed test runs, the live view greys out the ping lines and tags them
`[under hastighetstest]`. Those rows are kept out of the OK/DEGRADED/DOWN logic, so a
test never creates a false alarm. On exit you get two tables:

```
Target      Samples LossPct AvgMs MaxMs ... AvgMsUnderLoad MaxMsUnderLoad
192.168.2.1      10       0   0.6     3 ...              4             26
1.1.1.1          10       0   8.5    12 ...           12.4             34

Hastighet (Mbit/s):
Adapter    Tests Errors DownMedian DownMin DownMax UpMedian UpMin UpMax
Ethernet 3     3      2      241.1   241.1   241.1    221.6 203.2 223.7
Ethernet       3      2       48.1    48.1    48.1     18.9  18.9    19
```

(Real output from a test run with one NIC shaped to 50/20 Mbit/s. `Errors 2` = two
downloads got Cloudflare 429; their uploads were still measured.)

Files in `-LogDirectory` (default `C:\Temp\ConnMonitor`), one set per run:

| File | Content |
|---|---|
| `ConnQuality_<computer>_<yyyyMMdd_HHmm>.csv` | One row per target and sample, incl. `UnderLoad` |
| `ConnSpeed_<computer>_<yyyyMMdd_HHmm>.csv` | One row per NIC and speed test |
| `ConnEvents_<computer>_<yyyyMMdd_HHmm>.log` | Start/stop, status changes, speed results, summary |

Stop either script with **Ctrl+C** at any time - files are written as they go and the
summary is still printed.
