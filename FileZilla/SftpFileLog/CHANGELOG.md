# Changelog - FileZilla SFTP file-transfer log

All notable changes to `Export-SftpFileLog.ps1`, `Install-SftpFileLogTask.ps1` and
`transfers.html` are documented here.
Format based on [Keep a Changelog](https://keepachangelog.com/);
this tool follows [Semantic Versioning](https://semver.org/).

Git tags for this tool are prefixed `filezilla-sftp-log-v*`.

## [1.0.0] - 2026-10-06

First public release, generalised from a tool used in production on a
FileZilla Pro Enterprise Server 1.9.4.

### Added
- `Export-SftpFileLog.ps1`: daily CSVs of Upload / Download / Delete / Rename /
  CreateDir, with failed operations paired to their open/delete
  (`UploadFailed`, `DownloadFailed`, `DeleteFailed` + Error text).
- Incremental rebuild suitable for a 5-minute schedule; shared-read of the live
  log; atomic writes; Excel-lock tolerant; rotating run log.
- `transfers.html`: offline search page (filters, failures-only, session
  drill-down, CSV export, auto-refresh, Open folder / Open file).
- `Install-SftpFileLogTask.ps1`: SYSTEM scheduled task, no overlap, 10-min
  limit, prefers PowerShell 7, `-RunNow` / `-Remove` / `-WhatIf`.
- All paths default to the script's own folder; `-Folder`, `-LogDir`, `-OutDir`
  overrides.
