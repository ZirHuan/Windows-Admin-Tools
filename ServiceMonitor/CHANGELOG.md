# ServiceMonitor Changelog

All notable changes to ServiceMonitor.ps1 are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/); versioning follows [SemVer](https://semver.org/).

## [Unreleased]

### Fixed
- **`tests/Test-ServiceMonitor.ps1`: repaired the suite against the v1.3.0 config
  hardening (16 of 23 tests were failing on both PS 5.1 and PS 7).** The suite forced
  legacy flat-file mode by passing `-ConfigFile` at a path that deliberately did not
  exist - which v1.3.0 turned into a deliberate fatal error, so every affected test
  died at startup and its assertions matched the crash log instead of real output.
  Tests now run a **copy** of `ServiceMonitor.ps1` from an empty temp directory and
  omit `-ConfigFile` entirely, reaching legacy mode via the documented default path.
  This also stops the suite writing `ServiceMonitor.state.json` / log artefacts into
  the working tree, since those resolve against the script's own folder.
  The script itself is unchanged - the hardening was correct, the test technique was not.

### Added
- **`tests/Test-ServiceMonitor.ps1`: coverage for the explicit-missing-config fatal.**
  The v1.3.0 hardening shipped with no test of its own, which is why breaking the
  suite with it went unnoticed until CI. Now asserts exit 1, the specific error
  message, and that no legacy fallback occurred.

### Changed
- **`tests/Test-ServiceMonitor.ps1`: tightened the services-file-not-found assertion**
  from `not found` to `Services file not found`. The loose form also matched the
  benign `monitor-config.json not found` fallback warning that legacy mode always
  logs, so the test passed for the wrong reason (same defect class as b2c47d9).

## [1.3.0] - 2026-07-02

Hardening release driven by a three-way critical review (Claude, powershell-script-tester
agent, Gemini second pass) after the v1.2.9 SYSTEM-task fix, followed by an adversarial
verification pass on the new code itself (12 runtime probes on PS 7) whose findings are
folded in below.

### Added
- **`ServiceMonitor.ps1`: alert cooldown (`-AlertCooldownMinutes`, default 60).**
  Previously a service that stayed broken emailed every group on every 5-minute
  tick (2 mails/tick = ~576/day) until someone intervened - burying real alerts
  and risking the relay account being rate-limited. Repeat alerts for the same
  service+condition (down / restart failed / not found / disabled / invalid
  name) are now suppressed until the cooldown passes, tracked in
  `ServiceMonitor.state.json` next to the script. RECOVERED always sends and
  clears the service's cooldown, so a NEW failure after recovery alerts
  immediately. `0` disables the cooldown.
- **`ServiceMonitor.ps1`: run time budget (`-RunBudgetMinutes`, default 20).**
  With many simultaneously failed services the serial restart loops (~2.5-4 min
  each) could exceed the scheduled task's 30-minute execution limit and get
  killed mid-run, leaving later services unchecked and unalerted. Past the
  budget, remaining down services are still detected and alerted but restart
  attempts wait for the next run.
- **`ServiceMonitor.ps1`: script-scope `trap` for last-resort error visibility.**
  Any terminating error that escapes to script scope (bad JSON config,
  unwritable log, unexpected .NET failure) is appended to the log and written
  to the Application event log (EventId 3001) before exiting 1. Closes the
  "task result 1, empty log, no clue" failure class that v1.2.9 was debugged
  through.
- **`ServiceMonitor.ps1`: alert email on rejected (invalid) service names** -
  previously a garbage entry from the web UI only produced a log line nobody
  reads; now it alerts the configured group (cooldown-gated).
- **`New-CredStore.ps1`: prompts for SMTP port (default 587), SSL, and From
  address** instead of silently defaulting to 25 / off / empty. An empty From
  is now an error, and a `host:port` value in the server field is rejected -
  all three were real deployment traps (relay accepted mail then dropped it
  because the From domain was not authorized).

### Changed
- **`ServiceMonitor.ps1`: exit codes are now meaningful.** 0 = OK, 1 = fatal
  script error, 2 = ran fine but at least one service failed - so
  `LastTaskResult` distinguishes "the monitor crashed" from "a service is
  broken" (they were both 1 before, which cost hours of misdiagnosis).
- **`ServiceMonitor.ps1`: empty / all-paused config is no longer an error.**
  A fresh install (web installer creates a config with zero services) logged
  ERROR + event 3000 and exited 1 every 5 minutes until services were added;
  pausing everything during maintenance did the same. Now exits 0 with an
  INFO/WARNING line.
- **`ServiceMonitor.ps1`: explicitly passed `-ConfigFile` that does not exist
  is now a fatal error** instead of silently falling back to a possibly-stale
  legacy `services.txt` (which could monitor the wrong service set forever).
- **`ServiceMonitor.ps1`: `-TestEmail` exits 1 when the send fails** instead of
  always 0 - operators script against this to verify SMTP before scheduling.
- **`ServiceMonitor.ps1`: StopPending services are waited on (up to 60 s)
  before a start is issued.** Previously `sc start` fired instantly into a
  still-stopping service, failed 3 times, and raised a false RESTART FAILED
  alert during slow planned restarts.
- **`ServiceMonitor.ps1`: `$ScriptDir` resolution now fails hard** with a clear
  message when the script is not run from a file, instead of guessing
  `Get-Location` (which under SYSTEM is C:\Windows\system32 and produced a
  silently-empty monitor). Path params are also normalized to absolute paths
  so .NET file APIs cannot resolve them against the wrong working directory.
- **`Install-ScheduledTask.ps1`: task action now includes `-NoProfile`**
  (SYSTEM no longer loads the all-users profile every 5 minutes), and the
  exists-check/unregister are scoped to `-TaskFolder` so a same-named task in
  a different scheduler folder is not removed by mistake.
- **`New-CredStore.ps1`: requires elevation** (`#Requires -RunAsAdministrator`)
  - the smtp.key ACL restriction and the usual C:\ServiceMonitor target need
  it anyway; `Get-Acl`/`Set-Acl` use `-LiteralPath`. **BREAKING** for anyone
  who ran it non-elevated into a user-writable folder (portable use), and the
  directive also errors on non-Windows PowerShell - run tool-script CI checks
  on Windows runners.
- **`ServiceMonitor.ps1`: `-RunBudgetMinutes 0` disables the budget** instead
  of making the deadline "now" (which would have skipped every restart).
- **`Install-ScheduledTask.ps1`: `-TaskFolder` is normalized** to the
  `\Folder\` shape Task Scheduler expects, so `Monitoring` / `\Monitoring`
  no longer fail the exists-check and registration.
- **`New-CredStore.ps1`: the SMTP port prompt validates its input**
  (numeric, 1-65535) instead of crashing with a cast error or silently
  writing a junk port into smtp.json.
- **`Install-ScheduledTask.ps1`: task now runs on battery / DC power**
  (`-AllowStartIfOnBatteries -DontStopIfGoingOnBatteries`).
  `New-ScheduledTaskSettingsSet` defaults BOTH battery flags to blocking, so on
  a laptop - or a VM that reports its power source as DC - Task Scheduler parks
  the task in state `Queued` and never launches the process: `LastTaskResult`
  stays `0`, no log line is written, and nothing is monitored. Diagnosed live on
  a VM RDP host where a stopped service went unrestarted and unalerted for 2 h
  while the interactive script ran perfectly. This is the single most important
  fix in the release for anyone whose monitor host is virtual or a laptop.

### Fixed
- **`ServiceMonitor.ps1`: cooldown timestamp commits only after a successful
  send.** The first cut stamped the state file inside `Test-AlertDue` *before*
  the email went out, so a failed send (relay briefly down) silenced that
  alert for the whole cooldown window, and `-NoEmail` test runs poisoned the
  real task's state. Now `Test-AlertDue` is check-only and `Set-AlertSent`
  stamps after `Send-Alert` reports success - and never under `-NoEmail`.
- **`ServiceMonitor.ps1`: cooldown state file is written UTF-8, not ASCII.**
  A non-ASCII character in a state key (e.g. Swedish error text) would be
  mangled to `?` on save, never match on re-read, and re-alert every run -
  a permanent alert storm, the exact failure the cooldown exists to prevent.
- **`ServiceMonitor.ps1`: `-SmtpCredFile`/`-SmtpKeyFile`/`-SmtpConfigFile`
  are normalized to absolute paths** like the other path params (relative
  values would resolve against the process working directory in .NET calls).
- **`install-monitor-web.ps1`: Python version probes flattened with
  `Out-String`.** The `& python --version 2>&1` capture can be an array, and
  `-notmatch` on an array *filters* rather than tests - a multi-line result
  could skip a usable interpreter.
- **`install-monitor-web.ps1`: stray `confirm` argument removed from
  `nssm stop`** (only valid for `nssm remove`); the reinstall path now
  actually stops the old service instead of erroring silently.
- **`install-monitor-web.ps1`: `SecurityProtocol` is OR-ed with Tls12** in the
  python.org fallback instead of assigned, matching the NSSM download path -
  plain assignment clobbered already-enabled protocols (e.g. TLS 1.3) for the
  rest of the session.
- **`ServiceMonitor.ps1`: recipient dedup is now case-insensitive** (in
  `Send-Alert`, the single choke point) - `User@x.com` + `user@x.com` no
  longer get every alert twice, including across the merged 'both' group.
- **`ServiceMonitor.ps1`: alert bodies no longer leak cross-group addresses.**
  The embedded log tail filtered out `Alert sent to:` lines, which listed the
  OTHER group's recipient addresses in dev-only/iver-only alerts.
- **`ServiceMonitor.ps1`: log writability probe no longer appends a blank line
  every run** (288 blank lines/day); uses a zero-length .NET append instead.
- **`ServiceMonitor.ps1`: `sc.exe` output capture** moved under
  `$ErrorActionPreference='Continue'` - stderr output could become a
  terminating NativeCommandError before `$LASTEXITCODE` was read, misreporting
  a possibly-successful start as failed.
- **`ServiceMonitor.ps1`: service-name validation is now enforced inside
  `Get-ServiceStartupType` / `Invoke-ServiceRestart`** (safe by construction
  for the WQL filter and sc.exe interpolation), not only at the call site.
- **`ServiceMonitor.ps1`: misleading "Alerts suppressed (alerts=none)" log
  line** now reports the actual routing value and resolved recipient count, so
  "group has no recipients configured" is distinguishable from a deliberate
  `none`.
- **`install-monitor-web.ps1`: `Find-Python` PATH loop now verifies pip** (the
  fallback scan already did) and all native probes run under a function-local
  `$ErrorActionPreference='Continue'`, so a stderr warning can no longer skip
  a usable interpreter and trigger an unnecessary system-wide install.
- **`install-monitor-web.ps1`: `nssm install` / `nssm start` exit codes are
  checked** - a failed install (e.g. service "marked for deletion") no longer
  cascades into ten follow-up errors and a false "Service started" message.
- **All three tool scripts (`Install-ScheduledTask`, `New-CredStore`,
  `install-monitor-web`): `$PSScriptRoot` removed from param() defaults** -
  the same launch-shape bug fixed in ServiceMonitor.ps1 in v1.2.9 (empty under
  `powershell.exe -File`, e.g. Explorer's "Run with PowerShell").

## [1.2.9] - 2026-07-02

### Fixed
- **`ServiceMonitor.ps1`: scheduled task failed with exit 1 / no log under SYSTEM.**
  The task ran (`LastTaskResult: 1`) but produced no log and never restarted the
  monitored service, while an interactive `.\ServiceMonitor.ps1` run worked fine.
  Root cause: the `$ConfigFile`/`$ServicesFile`/`$RecipientsFile`/`$LogFile` param
  defaults called `Join-Path $PSScriptRoot '...'`, but `$PSScriptRoot` is **empty
  inside the `param()` block** when the script is launched via `powershell.exe
  -File ...` (how the SYSTEM scheduled task invokes it). The empty `-Path` made
  `Join-Path` throw at parameter binding - before `Initialize-Log`, so nothing was
  ever written. Relative-path invocation (`.\ServiceMonitor.ps1`) populates
  `$PSScriptRoot`, which is why manual runs masked the bug. The four params now
  default to `''` and the script directory is resolved *after* the param block via
  layered fallbacks (`$PSScriptRoot` -> `$PSCommandPath` -> `$MyInvocation` ->
  `Get-Location`), then any unset path is filled from it. The `smtp.json`
  auto-locate fallback uses the same resolved `$ScriptDir`.

## [1.2.8] - 2026-07-01

### Added
- **`monitor_web.py`: audit-log user now defaults to the logged-on RDP user.**
  The change log previously showed "unknown" (or nagged for a name) until a user
  typed one, because the service runs as LocalSystem and cannot read the RDP
  user's identity from the environment. It now resolves the interactive user
  behind each `127.0.0.1` request by mapping the client TCP port to the owning
  browser process (`GetExtendedTcpTable`), then that PID to its session
  (`ProcessIdToSessionId`) and the session to `DOMAIN\user`
  (`WTSQuerySessionInformationW`) - accurate even on a multi-user RDP host. Pure
  `ctypes`, no new dependencies; best-effort, so any failure falls back to the
  manual "Your name" field. An explicitly typed name still takes precedence.
- **`install-monitor-web.ps1`: post-install "best practice" guidance.** The
  completion summary now tells the operator to run `New-CredStore.ps1` and
  `Install-ScheduledTask.ps1` from the installed folder (`$InstallDir`), in order
  (cred store first), so the scheduled task's baked-in action and cred-file paths
  point at permanent copies rather than a temp/download folder that may be deleted.

## [1.2.7] - 2026-07-01

### Fixed
- **`install-monitor-web.ps1`: pip step aborted on a benign warning.** After the
  Python auto-install, `pip install` printed a harmless warning to stderr
  (`WARNING: The script idna.exe is installed in '...\Scripts' which is not on
  PATH`). Because `$ErrorActionPreference='Stop'` turns any native-command stderr
  captured via `2>&1` into a *terminating* NativeCommandError, the install died on
  the warning even though pip exited 0 and installed the packages. The pip call
  (and the winget install loop, which had the same latent pattern) now capture
  under `$ErrorActionPreference='Continue'` and decide success from the exit code.

## [1.2.6] - 2026-07-01

### Added
- **`install-monitor-web.ps1`: automatic all-users Python install.** A per-user
  Python (under `\Users\...`) is invisible to other accounts and to the SYSTEM
  service NSSM registers, so a fresh Administrator login previously failed with
  "Python 3 not found" even though a per-user Python existed under another
  profile. When no system-wide Python is found - or only a per-user one is - the
  installer now installs Python 3 for all users automatically: `winget install
  --scope machine` first (trying 3.13 / 3.12 / 3.14), falling back to the
  official python.org silent installer (`/quiet InstallAllUsers=1 PrependPath=1`),
  with the version resolved dynamically and HEAD-verified before download. Pass
  `-NoPythonAutoInstall` to disable and fail with manual instructions instead.

### Changed
- Python detection refactored into `Find-Python` (with a `-SystemOnly` mode used
  after an all-users install so a per-user copy still on PATH is not re-selected).

## [1.2.5] - 2026-06-29

### Fixed
- **HIGH** `Install-ScheduledTask.ps1`: scheduled-task registration failed at
  `Register-ScheduledTask` (line ~142). The `RepetitionDuration [TimeSpan]::MaxValue`
  trigger is accepted by `New-ScheduledTaskTrigger` but rejected at *registration*
  on some Windows builds, so the old try/catch (which wrapped only trigger
  creation) never fired and registration threw. Registration now retries: attempt
  `MaxValue` ("Indefinitely"), then fall back to a 10-year duration, throwing only
  if both fail (surfacing the underlying error).

## [1.2.4] - 2026-06-29

### Added
- **Windows Event Log integration** in `ServiceMonitor.ps1`. WARNING and ERROR
  messages (service down, not found, disabled, restart-attempt failures) are now
  mirrored to the **Application** event log under source `ServiceMonitor`
  (WARNING = EventId 2000, ERROR = 3000), and a successful restart writes an
  Information event (EventId 1001, "recovered"). INFO chatter stays file-only to
  keep the event log readable. The event source is created on first run (requires
  admin / the SYSTEM scheduled-task account); if it cannot be created, file
  logging continues and event writes are skipped gracefully.

## [1.2.3] - 2026-06-29

### Added
- Per-deployment SMTP settings via `smtp.json`. `New-CredStore.ps1` now prompts
  for (and stores) the SMTP **server**, plus port/from/useSsl, in `smtp.json`
  alongside `smtp.key`/`smtp.cred`/`smtp.user` - so the relay can differ per
  server instead of relying on a hardcoded default. `ServiceMonitor.ps1`
  auto-loads `smtp.json` (new `-SmtpConfigFile` param; otherwise looked up next to
  `-SmtpCredFile`, then beside the script). Explicit `-Smtp*` / `-FromAddress`
  parameters still override the file. New-CredStore bumped to 1.2.0.

### Fixed
- **HIGH** `install-monitor-web.ps1`: also install `python-multipart`. The web UI
  uses FastAPI `Form()` (the `/auth` token route); without `python-multipart`
  uvicorn raises `RuntimeError: Form data requires "python-multipart"` at startup
  and never binds, so the service sat in NSSM's Paused (throttled) state and the
  page was unreachable. Installer previously only installed `fastapi` + `uvicorn`.

## [1.2.2] - 2026-06-29

### Fixed (Windows host testing, 2026-06-29)
- **HIGH** `install-monitor-web.ps1`: on a re-run, Python and NSSM detection now
  succeed (per 1.2.1), but the file-copy step then threw "Cannot overwrite the
  item ...\nssm.exe with itself" because nssm.exe was detected *inside*
  `$InstallDir` and then copied onto itself. The copy is now skipped when the
  resolved source and destination paths are equal.

## [1.2.1] - 2026-06-29

### Fixed (powershell-script-tester review, 2026-06-26)
- **CRITICAL** `install-monitor-web.ps1`: force TLS 1.2 before downloading NSSM from nssm.cc.
  Windows PowerShell 5.1 defaults to SSL3/TLS1.0 and the download would otherwise fail.
- **CRITICAL** JSON BOM mismatch: `install-monitor-web.ps1` now writes `monitor-config.json`
  without a UTF-8 BOM (`Write-JsonFile` helper using `UTF8Encoding($false)`), and `monitor_web.py`
  reads config with `utf-8-sig` as belt-and-braces. `Set-Content -Encoding UTF8` on PS 5.1 emitted
  a BOM that crashed Python's `json.load`.
- **HIGH** `ServiceMonitor.ps1`: `ConvertFrom-LegacyServices` (was `Parse-LegacyServices`) now returns
  a single object `@{Active;Paused}` instead of `$a, $b`. An empty `Active` array previously vanished
  in the pipeline, promoting paused services to active in all-paused legacy configs.
- **HIGH** `Install-ScheduledTask.ps1`: added `-SmtpUser/-SmtpCredFile/-SmtpKeyFile/-SmtpUseSsl`
  pass-through into the task action so authenticated relay actually runs under SYSTEM (was dead config).
- **MED** `tests/Test-ServiceMonitor.ps1`: healthy-service test now ignores the expected
  legacy "config not found" warning; all-paused test fixed by the `ConvertFrom-LegacyServices` change.
- **MED** `monitor_web.py`: optional `SM_TOKEN` access-token gate (default off; RDP login remains the
  primary auth). `install-monitor-web.ps1 -AccessToken` wires it into the NSSM service environment.
- **LOW** Hardening: StrictMode-safe nested JSON reads (`Get-Prop`); `RepetitionDuration`
  `[TimeSpan]::MaxValue` now falls back to a 10-year span if rejected; `New-CredStore.ps1` frees BSTR
  buffers via `ZeroFreeBSTR` and uses `RandomNumberGenerator::Create()` (RNGCryptoServiceProvider is
  obsolete); `Send-Alert` disposes SmtpClient/MailMessage in `finally`; renamed unapproved verbs
  (`Load-Config`→`Import-MonitorConfig`, `Build-AlertBody`→`New-AlertBody`); desktop shortcut is now a
  proper `.url` internet shortcut instead of a `.lnk`.

### Fixed (Windows host testing, 2026-06-29)
- **HIGH** `install-monitor-web.ps1`: Python detection failed on a host where PATH
  exposed only the Windows Store stubs (`py.exe`/`python.exe`/`python3.exe` under
  `WindowsApps`). These are correctly skipped (they do not work under a service
  account), but the script then threw without checking whether a real Python was
  installed elsewhere. Added a fallback that scans the registry
  (`HKLM`/`HKCU` `Python\PythonCore\*\InstallPath`, incl. WOW6432Node) and standard
  filesystem locations (`Program Files`, `Program Files (x86)`, per-user
  `LOCALAPPDATA\Programs\Python`, `C:\Python3*`) before giving up. System-wide
  installs are preferred over per-user. Each candidate is validated for both a 3.x
  version and a functional `pip` (a Python with broken pip is skipped). The final
  error now also tells the user to disable the Store App execution aliases.
- **HIGH** `install-monitor-web.ps1`: NSSM acquisition was a single hardcoded
  `nssm.cc` download with no retry, so it died when the site returned 503. Now
  retries the stable + CI URLs, clears stale/partial temp artifacts between
  attempts, and falls back to `winget install NSSM.NSSM` (which uses its own CDN
  and verifies the package, so it works when nssm.cc is down). After winget it
  resolves the binary via `Get-Command` first, then the package store. Final error
  points to `winget install NSSM.NSSM` and the existing `-NssmPath` parameter.
- **HIGH** `ServiceMonitor.ps1`: fixed StrictMode (`-Version Latest`) crashes that
  failed ~7 Pester tests on the Windows CI runner. (1) `Get-AlertRecipients`
  returns an unrolled empty array (becomes `$null`), so the five `$alertsTo.Count`
  checks threw "property 'Count' cannot be found"; `$alertsTo` is now wrapped in
  `@()` at assignment. (2) The "Active services to check" / "Paused" log lines used
  `.Name` member access and `.Count` on possibly-empty collections; now built via
  `ForEach-Object` with `@()`-wrapped counts. (3) `$tail` in `New-AlertBody`
  wrapped in `@()`. (4) The `-TestEmail` short-circuit now runs *before* the
  "Active services to check" log so TestEmail mode no longer logs it.
- `tests/Test-ServiceMonitor.ps1`: the recipients fixture now includes a real
  address so the `[NoEmail]` suppression path is actually exercised.

## [1.2.0] - 2026-06-26

### Added
- `monitor_web.py`: FastAPI web admin UI (localhost:8080, no auth needed — RDP session is the
  auth layer). Two-panel layout: monitored services + available services. Mail groups panel with
  add/remove per group. Audit log panel (last 20 changes). Username cookie for audit attribution.
  One batched PowerShell call per page load for service status. Atomic config writes (temp→rename).
  Available services panel excludes ~80 standard Windows system services automatically.
- `install-monitor-web.ps1`: installs Python deps (fastapi+uvicorn), copies files, registers the
  web app as a Windows service via NSSM, places a desktop shortcut for all RDP users. Includes
  automatic migration of existing services.txt + recipients.txt into monitor-config.json.
- `monitor-config.json`: unified config format replacing services.txt + recipients.txt. Combines
  services (with per-service alert routing) and two mail groups (dev + iver).

### Changed
- `ServiceMonitor.ps1` v1.2.0: reads `monitor-config.json` when present (JSON mode). Each service
  entry carries an `alerts` field: `both` | `dev` | `iver` | `none` — controls which mail group
  receives alerts for that service. `none` fully suppresses alerts for that service. Legacy
  services.txt + recipients.txt fallback preserved for backward compatibility when the JSON file
  is absent. `Build-AlertBody` now includes `Sent to: <group>` line so recipients know why they
  received the alert.
- Alert emails now show which group the alert was routed to (Dev Team only / Iver Support only /
  Both groups) in the email body.

## [1.1.0] - 2026-06-26

### Added
- `SmtpUser`, `SmtpCredFile`, `SmtpKeyFile` parameters: support for authenticated SMTP relay.
  Credentials are stored as an AES-256 encrypted file pair created by `New-CredStore.ps1`.
  `[System.IO.File]::ReadAllBytes` used for key loading (PS 5.1 and 7+ compatible — avoids
  `-Encoding Byte` vs `-AsByteStream` difference).
- `SmtpUseSsl` switch: enables SSL/TLS on the SMTP connection.
- `-TestEmail` switch: sends a test alert to all recipients and exits without checking services.
  Use to verify SMTP connectivity and credentials after initial setup.
- `New-CredStore.ps1`: interactive credential setup helper. Generates AES-256 key with
  `RNGCryptoServiceProvider`, encrypts SMTP password with `ConvertFrom-SecureString -Key`,
  restricts key file ACL to SYSTEM + Administrators, and performs a round-trip verification.
- `Build-AlertBody`: appends last 20 lines of the log file to all alert emails for context.
- `tests/Test-ServiceMonitor.ps1`: Pester v5 integration test suite (25 tests across 9 `Describe`
  blocks). All tests use `-NoEmail` — no real SMTP calls. Covers: banner, running/paused/missing
  services, injection prevention, empty file, log rotation, `-TestEmail` mode, credential
  parameter validation, summary line, and exit codes.
- `tests/test-services.txt` / `tests/test-recipients.txt`: test data files.
- `ServiceMonitor-TestSet-v1.1.0.zip`: distributable package of all files.

### Changed
- `Initialize-SmtpCredential`: new function loads and decrypts SMTP credentials at startup;
  warns and disables auth (anonymous relay) if files are missing or decryption fails.
- `Send-Alert`: applies `$smtp.Credentials` and `$smtp.EnableSsl` when credentials are loaded;
  anonymous relay continues to work when no credential files are configured.
- Exit code: exits 1 when `$failureCount > 0` (permanent failures after all retries exhausted).
  Previously always exited 0; Task Scheduler can now detect and report failed runs.
- `Install-ScheduledTask.ps1`: version string in task description is now read dynamically from
  `ServiceMonitor.ps1` (`$ScriptVersion` pattern match) — no longer hardcoded.
- `Install-ScheduledTask.ps1`: added `-PowerShell7` switch to use `pwsh.exe` instead of
  `powershell.exe`; resolves default PS7 install path and falls back to `pwsh.exe` on PATH.

## [1.0.1] - 2026-06-24

### Added
- `StartSettleSeconds` parameter (default 8 s): always waits after `sc.exe start`
  before re-checking status - prevents false "RESTART FAILED" for slow-starting services.
- Log rotation: renames log to `.1`/`.2`/`.3` when it exceeds `LogMaxSizeMB` (default 10 MB).
- Log directory auto-creation at startup; hard-fails if log is unwritable.
- Service-name validation against `^[A-Za-z0-9_.$ -]+$` before passing to `sc.exe`
  (prevents argument injection from a tampered services.txt).
- `StartPending` state handling: waits up to 30 s for a transitioning service
  before declaring it DOWN and triggering a restart.
- `Disabled` startup-type detection: reports and skips rather than looping 3 restarts.
- `Wait-ServiceState` helper: polls until a service reaches the target state or times out.
- `Get-ServiceStartupType` helper via `Win32_Service` CIM class.
- `[System.IO.File]::ReadAllLines` for config files: handles UTF-8 BOM transparently.
- Recipients and service list always coerced to arrays (`@(...)`) - fixes StrictMode
  failure on single-item or empty results.
- `Install-ScheduledTask.ps1`: added `-RepetitionDuration ([TimeSpan]::MaxValue)` so
  the task repeats indefinitely (not just until an implicit deadline).
- `Install-ScheduledTask.ps1`: fixed `ExecutionTimeLimit` to a fixed 30-minute safe
  value independent of the polling interval; validates `IntervalMinutes >= 1`.

### Changed
- `Invoke-ServiceRestart`: waits for service to reach `Stopped` state (via `WaitForStatus`)
  before issuing start - prevents start/stop race on hung services.
- Status re-check after each restart attempt now separated into: settle wait first,
  then inter-attempt delay if more attempts remain (was: inter-attempt delay first).
- `New-Object` used instead of `::new()` for SmtpClient and MailMessage for clarity
  (both work on PS 5.1+; `New-Object` is more idiomatic for 5.1 scripts).
- Exit code is always 0 on a clean run; failure count is in the log and email, not the exit code.

## [1.0.0] - 2026-06-24 (initial, superseded)

Initial release. Not deployed.
