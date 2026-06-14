# Changelog - Org-wide Calendar Sharing scripts

All notable changes to `Set-HistoricalCalendarPrivate.ps1` and
`Enable-OrgCalendarSharing.ps1` are documented here. These two scripts ship as a
pair and are versioned together.

Format based on [Keep a Changelog](https://keepachangelog.com/);
follows [Semantic Versioning](https://semver.org/).

Git tags for this tool are prefixed `calendar-sharing-v*` (this repo holds
multiple independent scripts).

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
