# Changelog - AD

All notable changes to the scripts in `AD/`. Each script is versioned on its own.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
Git tags are prefixed with the tool name (for example `fsmo-role-holder-v1.0.0`).

## Building blocks (new 2026-11-19)

Small read-only scripts that each do one thing, run on their own and are called by the
larger jobs (`DomainScanner/Invoke-DomainHealthAudit.ps1`, `AD/Manage-FSMORoles.ps1`)
instead of having the same code inside each of them.

### Get-FsmoRoleHolder.ps1 [1.0.0] - 2026-11-19
- The five FSMO roles and their holders, one object per role, with `-ExpectedHolder`
  (short-name compare). Role ids match `Move-ADDirectoryServerOperationMasterRole`.

### Get-PrivilegedAccountRisk.ps1 [1.0.0] - 2026-11-19
- Every user in the AdminSDHolder-protected groups, scored OK / REVIEW / HIGH-RISK.
- Groups by well-known SID (works on non-English DCs); forest-root groups looked up in
  the root domain. Unexpandable groups are non-terminating errors (`-ErrorVariable`).

### Get-UnsignedLdapClient.ps1 [1.0.0] - 2026-11-19
- Clients doing unsigned / simple LDAP binds per DC (events 2889/2887), with the
  logging level and a status that says whether the result can be trusted.
  `-ShowEnableCommand` prints how to turn per-client logging on.

All three tested on a lab domain (Server 2022, Windows PowerShell 5.1), 2026-11-19.

## Manage-FSMORoles.ps1

### [1.2.0] - 2026-11-19
#### Changed
- Moved from the repo root (`Manage-FSMORoles_0.2.ps1`) to `AD/Manage-FSMORoles.ps1`.
- Role listing and post-move verification use `Get-FsmoRoleHolder.ps1`; LDAP/Kerberos
  reachability uses `Network/Test-TcpPort.ps1` instead of `Test-NetConnection`.
- Post-move check compares short names exactly (`DC1` no longer matches `DC10`).
#### Fixed
- Single-DC domains stopped at step 2 on Windows PowerShell 5.1 ("The property 'Count'
  cannot be found"): one DC came back as a plain string under `Set-StrictMode -Version
  Latest`. Same fix for the health-score count.
#### Tested
- Lab domain (one DC), PS 5.1, all eight steps including the (no-op) move and the
  post-move verification. Health score 8/8.

## Get-ADGroupMembers.ps1

### [1.1.0] - 2024-12-11

#### Added
- Loop functionality: Press ENTER to run again or ESC to exit
- Screen clearing between runs for better user experience
- Improved exit message

#### Changed
- Script now runs continuously until ESC is pressed
- Enhanced user prompts and instructions

### [1.0.0] - 2024-12-11

#### Added
- Initial release
- Interactive numbered list of AD groups
- Group member display functionality
- Color-coded output
- Input validation
- Support for quitting with 'Q'
- Display of member object types (user, computer, group)
- Total member count display
- Handling of empty groups

#### Fixed
- Replaced problematic ReadKey implementation for better compatibility