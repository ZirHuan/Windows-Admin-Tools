# Changelog - Network

Format based on [Keep a Changelog](https://keepachangelog.com/). Tags:
`wifi-connection-vX.Y.Z` (Test-WifiConnection) and `watch-connection-quality-vX.Y.Z`
(Watch-ConnectionQuality). Usage and examples: [README.md](README.md).

## Test-WifiConnection.ps1

### 1.1.0 - 2026-11-04

- New `-SpeedTestOnly` switch: runs the per-NIC speed test (download + upload) once
  on every selected adapter, one at a time, prints a result table and exits. No
  latency measurement and no latency CSV; results go to `-LogPath` (default
  `Desktop\netspeed-<date>-<time>.csv`). `-AdapterName`, `-SpeedTestSec`,
  `-SpeedTestStreams` and `-SpeedTestHost` apply as before; `-DurationSec`,
  `-IntervalMs` and `-SpeedTestIntervalMin` are ignored.
- The same speed-test engine is now also used by `Watch-ConnectionQuality.ps1`
  v1.2.0 (repository root). The C# block is an identical copy with the same type
  names, so both scripts can run in one PowerShell session - change both together.

Verified on Windows Server 2022 with two NICs, the second shaped to 50/20 Mbit/s on
the hypervisor (tbf egress + ingress police):
- PS 5.1.20348: both NICs in one run - unshaped 231.3 down / 221.8 up, shaped
  48.1 / 19.0, so each NIC is measured over its own link.
- PS 7.4.5: `-AdapterName 'Ethernet'` tested only the matching NIC; `-Delimiter ';'`
  honoured. Download hit Cloudflare HTTP 429 after the day's earlier tests; it was
  reported as an error for that direction and upload (19.4) was still measured.
- Simulated Ctrl+C mid-test: header shows `AVBRUTET`, CSV closed with header only,
  pipeline stopped in 16 ms.

### 1.0.0 - 2026-11-03

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

## Watch-ConnectionQuality.ps1 (repository root)

### 1.2.0 - 2026-11-04

- **Per-NIC speed test** (download + upload over HTTPS to `speed.cloudflare.com`),
  10 s after start and then every `-SpeedTestIntervalMin` minutes (default 180,
  `0` = off - the v1.1 behaviour). `-SpeedTestAdapter` selects the NICs (wildcards;
  default all active NICs with IPv4 and a default gateway). `-SpeedTestSec`,
  `-SpeedTestStreams` and `-SpeedTestHost` as in Test-WifiConnection. NICs are tested
  one at a time with TCP streams bound to each NIC's address.
- Results go to a new `ConnSpeed_<computer>_<yyyyMMdd_HHmm>.csv` and the event log;
  the exit summary gets a per-NIC table (tests, errors, median/min/max down and up).
- Pings taken while a speed test runs get `UnderLoad = True` in the CSV, are shown
  greyed out with `[under hastighetstest]`, and are kept out of status, events and the
  normal statistics so a test never raises a false DEGRADED. They are reported
  separately as `AvgMsUnderLoad` / `MaxMsUnderLoad` (bufferbloat).
- A test still running when the duration ends is allowed to finish; Ctrl+C cancels it
  and logs a warning.
- **Behaviour change:** speed tests are on by default and consume data (roughly line
  rate x `SpeedTestSec` per direction and NIC). Use `-SpeedTestIntervalMin 0` to keep
  the old behaviour.
- Source converted to pure ASCII (a/o for the Swedish letters), as required by the
  repository encoding convention; summary tables in the event log are no longer
  wrapped at console width.

Verified on Windows Server 2022, PS 5.1.20348, two NICs (one shaped to 50/20 Mbit/s):
first test 241.1 / 223.7 and 48.1 / 18.9 Mbit/s; 18 of 48 ping rows flagged
`UnderLoad`; later tests hit Cloudflare HTTP 429 on download, logged as WARN while
monitoring continued. Simulated Ctrl+C during a test logged
`Pagaende hastighetstest avbrutet.` and wrote the summary. `-SpeedTestIntervalMin 0`
tested in a PS 7 container (no speed CSV, unchanged summary).

### 1.1.1 - 2026-11-03

- Removed a stray pasted line from the help text.

### 1.1.0 - 2026-11-03

- Log files named after the start time (`yyyyMMdd_HHmm`), one set per run; version
  written to the event log at start.

### 1.0.0 - 2026-11-03

- First version: latency, jitter and packet loss per target, default-gateway
  detection, event log on status change, summary, PS 5.1 / 7 compatible.
