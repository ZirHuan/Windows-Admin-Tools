# Changelog - repository-wide

Notable changes that cut across the whole repository. Individual tools keep their
own changelogs and their own version numbers:

- [`AD/CHANGELOG.md`](AD/CHANGELOG.md)
- [`M365/README.md`](M365/README.md) (Get-CustomerReport changelog in the script's `.NOTES`)
- [`M365/CHANGELOG-CalendarSharing.md`](M365/CHANGELOG-CalendarSharing.md)
- [`Mail/CHANGELOG.md`](Mail/CHANGELOG.md)
- [`DomainScanner/CHANGELOG.md`](DomainScanner/CHANGELOG.md)
- [`ServiceMonitor/CHANGELOG.md`](ServiceMonitor/CHANGELOG.md)
- [`Network/CHANGELOG.md`](Network/CHANGELOG.md)
- [`FileZilla/SftpFileLog/CHANGELOG.md`](FileZilla/SftpFileLog/CHANGELOG.md)

Format based on [Keep a Changelog](https://keepachangelog.com/).
Git tags are per-tool and prefixed with the tool name (this repo holds multiple
independent scripts), so a repository-wide entry like the one below is not tagged.

## 2026-11-19

### Changed - duplicates removed, shared code broken out into building blocks

A review of the whole repository found the same tools and the same code in several
places. Each is now in one place; the larger jobs call small, separately runnable
scripts instead of carrying their own copy.

**Duplicates merged or removed**

| Was | Now |
|---|---|
| `AD/Inventering/Invoke-DomainHealthAudit.ps1` + `Get-CertBlastRadius.ps1` (based on the buggy 1.0.0 seed, with extra checks) and `DomainScanner/` 1.1.0 (with the fixes) | `DomainScanner/` only; 1.2.0 carries the extra checks, rewritten |
| `M365/Get-TenantOverview.ps1` | Removed - a strict subset of `Get-CustomerReport.ps1` |
| `Test-OpenPorts.ps1` (root, hard-coded customer IP) and the TCP probe inside the domain audit | `Network/Test-TcpPort.ps1` |
| 297-line C# engine + NIC selection copied in `Test-WifiConnection` and `Watch-ConnectionQuality` | `Network/NetProbe.Engine.ps1` |
| `customers.json` lookup written three ways in the M365 scripts | `M365/Resolve-M365Customer.ps1` |
| `Manage-FSMORoles_0.2.ps1` (root) | `AD/Manage-FSMORoles.ps1` |

**New building blocks** (each runs on its own and is called by a larger job):
`AD/Get-FsmoRoleHolder.ps1`, `AD/Get-PrivilegedAccountRisk.ps1`,
`AD/Get-UnsignedLdapClient.ps1`, `Network/Test-TcpPort.ps1`,
`Network/Get-TlsCertificate.ps1`, `Network/Invoke-NicSpeedTest.ps1`,
`M365/Resolve-M365Customer.ps1`.

**Version bumps**: Invoke-DomainHealthAudit 1.2.0, Manage-FSMORoles 1.2.0,
Test-WifiConnection 1.2.0, Watch-ConnectionQuality 1.3.1, Get-CustomerReport 1.6.1,
move365user-shared 1.1.0, Get-CADeletionAudit 1.1.0. Details in each folder's changelog.

**Bugs found while testing** (fixed): `Manage-FSMORoles` stopped on single-DC domains
on PS 5.1; `move365user-shared` read `DefaultLicenseSku` while the template and
`Get-CustomerReport` write `PrimaryLicenseSku`, so the profile's license was never used.

**Deploying**: a script that uses building blocks needs them next to it or in the
sibling repo folder - copy the whole folder (or the release ZIP), not one file. A
missing block stops the script with a message naming it.

Tested on Windows Server 2022 / Windows PowerShell 5.1 in a lab domain and parsed on
PowerShell 7.4; see the folder changelogs. M365 scripts were not run against a tenant -
only `Resolve-M365Customer.ps1` was tested (offline, PS 5.1 + 7.4).

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
