# Changelog - Network

Format based on [Keep a Changelog](https://keepachangelog.com/). Tags: `wifi-connection-vX.Y.Z`.

## Test-WifiConnection.ps1

### 1.0.0 - 2026-11-03

First tracked version, rebuilt from an untracked draft.

- Tests **all active NICs** with IPv4 and a default gateway at the same time
  (`-AdapterName` accepts names and wildcards to narrow it down).
- ICMP via `IcmpSendEcho2Ex` (iphlpapi.dll) with the source address bound, run in
  parallel on dedicated threads - replaces serial `ping.exe -S` calls and
  locale-dependent text parsing.
- Fixed time slots (no drift), timestamp taken at the start of each measurement,
  overruns counted, exact number of measurements per run.
- Roaming window measured in seconds, not rows.
- `netsh wlan show interfaces` blocks matched to the adapter by GUID (multiple
  Wi-Fi adapters, English and Swedish Windows); RSSI and band recorded when shown.
- Gateway taken from the adapter's own default route (lowest metric).
- CSV written incrementally (survives Ctrl+C), delimiter follows the Windows list
  separator so it opens directly in Excel.
- Per-NIC summary (loss, median/p95/max) and cross-NIC spike correlation.
- Periodic **speed test per NIC** (`-SpeedTestIntervalMin`, default 180 = every 3 h,
  first run 10 s after start; 0 = off): download + upload over HTTPS to
  `speed.cloudflare.com`, 4 parallel TCP streams bound to the NIC's address,
  8 s per direction after a slow-start warm-up, NICs tested one at a time.
  Runs in the background; latency rows during the test are flagged `UnderLast`,
  excluded from spike stats and reported as latency under load (bufferbloat).
  Results in `<log>-hastighet.csv`.

Verified on Windows Server 2022 (PS 5.1.20348 and 7.4.5): parse 0 errors, netsh
parser 22/22 against sample output, two-NIC run with netem fault injection
(delay/loss attributed to the right NIC), edge cases and interrupt handling.
Speed test: NIC shaped to 50/20 Mbit/s on the host measured 48.1/19.4; unshaped NIC
230-234 down / 210-226 up against a source-bound `curl` reference of 205-232 down.
Cloudflare limits seen in testing: HTTP 403 for >= 100 MB per request (script uses
50 MB), HTTP 429 after roughly 20 tests in 1.5 h.
The Wi-Fi fields are verified only against sample `netsh` output - not yet
against a real wireless adapter.
