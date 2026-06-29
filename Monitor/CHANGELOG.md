# ServiceMonitor Changelog

All notable changes to ServiceMonitor.ps1 are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/); versioning follows [SemVer](https://semver.org/).

## [Unreleased]

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
