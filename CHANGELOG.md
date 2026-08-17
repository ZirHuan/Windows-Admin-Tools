# Changelog - repository-wide

Notable changes that cut across the whole repository. Individual tools keep their
own changelogs and their own version numbers:

- [`AD/CHANGELOG.md`](AD/CHANGELOG.md)
- [`M365/CHANGELOG-CalendarSharing.md`](M365/CHANGELOG-CalendarSharing.md)
- [`Mail/CHANGELOG.md`](Mail/CHANGELOG.md)
- [`ServiceMonitor/CHANGELOG.md`](ServiceMonitor/CHANGELOG.md)

Format based on [Keep a Changelog](https://keepachangelog.com/).
Git tags are per-tool and prefixed with the tool name (this repo holds multiple
independent scripts), so a repository-wide entry like the one below is not tagged.

## 2026-08-17

### Fixed

Six scripts were saved as UTF-8 **without** a BOM while containing non-ASCII
characters. Windows PowerShell 5.1 falls back to the legacy ANSI codepage for a
BOM-less `.ps1`, so every multi-byte UTF-8 sequence was decoded as several junk
characters. That desynchronises the tokenizer and breaks string terminators and
brace matching, producing **87 parse errors** in total. All six scripts refused to
run on Windows PowerShell 5.1.

Every affected file was re-saved as **UTF-8 with BOM**. No script content was
changed and there is no functional change.

| Script | Parse errors before | Non-ASCII chars |
|---|---:|---:|
| `M365/Get-CustomerReport.ps1` | 56 | 1200 |
| `Manage-FSMORoles_0.2.ps1` | 14 | 1749 |
| `M365/Get-TenantOverview.ps1` | 7 | 536 |
| `Watch-Memory.ps1` | 4 | 1 |
| `M365/Get-CADeletionAudit.ps1` | 3 | 151 |
| `M365/Get-LicensedUsers.ps1` | 3 | 6 |

After the fix all six parse with **0 errors**.

The errors were badly misleading, which is why this went unnoticed for so long:
they are reported on later lines that are themselves perfectly valid. The clearest
case is `Watch-Memory.ps1`, which contains exactly **one** em-dash, on line 18, and
failed with three errors all pointing at **line 31** - a pure-ASCII line. PowerShell
7 defaults to UTF-8 and parses all of these files fine, so the bug only ever
appeared on a 5.1 host.

Version bumps for the two scripts that carry a version:

- `M365/Get-CustomerReport.ps1` 1.5.3 -> **1.5.4**
- `Manage-FSMORoles_0.2.ps1` 1.1 -> **1.1.1**

The other four scripts carry no version header and were left unversioned.

### Notes

Verified on Windows Server 2022 (build 20348), Windows PowerShell 5.1, by parsing
every file with `[System.Management.Automation.Language.Parser]::ParseFile()`
before and after. `M365/move365user-shared.ps1` is the control case: it holds 22
non-ASCII characters but already had a BOM, and has always parsed cleanly.

Remaining non-ASCII characters were deliberately left in place. With a BOM present
they are harmless, and several are intentional output (box-drawing separators,
warning symbols) whose replacement would change rendered report and console output.
