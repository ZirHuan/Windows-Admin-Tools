# ServiceMonitor Changelog

All notable changes to ServiceMonitor.ps1 are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/); versioning follows [SemVer](https://semver.org/).

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
