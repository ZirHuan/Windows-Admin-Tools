# Changelog - Calendar Sharing scripts

Scripts in this family:
- `Set-HistoricalCalendarPrivate.ps1` — bulk-mark historical events as Private
- `Enable-OrgCalendarSharing.ps1` — org-wide Default permission on all calendars
- `New-CalendarSharingGroups.ps1` — create/maintain sharing groups
- `Sync-CalendarSharing.ps1` — reconcile per-user permissions from group membership

Format based on [Keep a Changelog](https://keepachangelog.com/);
follows [Semantic Versioning](https://semver.org/).

Git tags for this tool are prefixed `calendar-sharing-v*` (this repo holds
multiple independent scripts).

---

## [Enable-OrgCalendarSharing 1.2.0 / New-CalendarSharingGroups 1.1.0] - 2026-11-02

### Fixed
- `Enable-OrgCalendarSharing`: resolve calendar folder via `Get-MailboxFolderStatistics`
  (`FolderType=Calendar`) instead of hardcoded `:\Calendar` — fixes Swedish tenants
  where the folder is named `Kalender`.
- `Enable-OrgCalendarSharing`: removed `-SharingPermissionFlags None` — parameter is not
  supported on `Default`/`Anonymous` entries in Exchange Online (was causing all calls to fail).

### Added
- `New-CalendarSharingGroups` v1.1.0: added `philip.vitzthum@antilopgroup.se` to `Kal-Se-Antilop`.
- `Sync-CalendarSharing` v1.0.1: initial commit of group-based reconcile script.
  Members of `Kal-Se-Antilop` get `LimitedDetails` on all `Kal-Delad-Antilop` calendars;
  stale entries are cleaned on next run.

---

## [Set-HistoricalCalendarPrivate 1.2.0] - 2026-06-22

### Added
- App-only authentication: new `-TenantId`, `-ClientId`, `-ClientSecret` parameters enable
  client credentials flow (`Connect-MgGraph -ClientSecretCredential`). Required when the
  connecting account has no mailbox/license (e.g. unlicensed admin-only accounts).

### Changed
- Calendar event operations now use `Invoke-MgGraphRequest` (raw Graph HTTP) instead of
  `Get-MgUserEvent` / `Update-MgUserEvent`. Eliminates the `Microsoft.Graph.Calendar`
  module dependency — only `Microsoft.Graph.Authentication` + `Microsoft.Graph.Users`
  are now required.
- Added "Found: N event(s) before cutoff" progress line per mailbox.

---

## [1.0.0] - 2026-06-14

First versioned release. Solves: "let everyone in the tenant see everyone's
calendar, without exposing sensitive historical meetings — only going forward."

### Added
- `Set-HistoricalCalendarPrivate.ps1` — bulk-sets event `sensitivity = private`
  via Microsoft Graph (`Update-MgUserEvent`) on every event ending before a
  `-CutoffUtc` date, across all or selected mailboxes.
  - `-WhatIf` / `ShouldProcess` support for safe dry runs.
  - `-SkipRecurringMasters` to leave recurring series with future occurrences
    untouched (off by default = recurring series masked too).
  - `-ThrottleDelayMs` pause between PATCH calls to ease Graph throttling.
  - Skips events already marked Private; per-mailbox masked count + grand total.
- `Enable-OrgCalendarSharing.ps1` — sets the `Default` (everyone-in-org)
  permission on each Calendar folder via `Set-MailboxFolderPermission`.
  - `-AccessRight` of `AvailabilityOnly` | `LimitedDetails` (default) | `Reviewer`.
  - Explicitly sets `-SharingPermissionFlags None` so `ViewPrivateItems` is never
    granted — Private items stay masked at every access level.
  - `-WhatIf` / `ShouldProcess` support; per-mailbox result + total.

### Notes
- "Private" is an honour-based Outlook/EXO flag, **not** encryption. Full Access,
  Delegate + ViewPrivateItems, or an app with `Calendars.Read` can still read
  Private items. Use Microsoft Purview sensitivity labels with encryption for
  genuinely confidential meetings.
- Localised tenants: the Calendar folder may be named in the mailbox language
  (e.g. "Kalender"). See the note at the foot of `Enable-OrgCalendarSharing.ps1`.
- Requirements: PowerShell 7+; `Microsoft.Graph` (`Calendars.ReadWrite`) for the
  masking script; `ExchangeOnlineManagement` + Exchange admin role for the
  sharing script.
