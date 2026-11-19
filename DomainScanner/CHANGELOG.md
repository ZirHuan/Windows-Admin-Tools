# Changelog - DomainScanner

Two independently versioned read-only tools. Format based on
[Keep a Changelog](https://keepachangelog.com/); versions follow
[Semantic Versioning](https://semver.org/). Git tags are per tool:
`domain-health-audit-vX.Y.Z` and `cert-blast-radius-vX.Y.Z`.

The version also lives in each script's `.NOTES` block and in `$ScriptVersion`,
which is printed in the console banner and stamped into the HTML/CSV output and
the returned object - so every report says which build produced it.

## Invoke-DomainHealthAudit.ps1

### [1.1.0] - 2026-09-28

Reviewed and test-run in an isolated lab domain
(Server 2022, Windows PowerShell 5.1), both healthy and with a simulated dead CA
(CA service stopped + a linked autoenrollment GPO).

#### Fixed
- **Did not parse anywhere.** Three `"...$dc: $_"` strings are invalid variable
  references (3 parse errors on PS 7), and non-ASCII characters without a BOM gave
  20 parse errors on PS 5.1. Now pure ASCII + UTF-8 BOM, 0 parse errors on both.
- **`-Html` never worked** - the local `$html` variable is the same variable as the
  `[switch]$Html` parameter (names are case-insensitive), so the page failed to
  assign and the report file contained just `True`.
- **TCP probe reported refused ports as open.** `WaitOne()` also returns true when the
  connect fails fast; `EndConnect()` is now called. Affected NET ports, stale-IP and
  LDAPS checks.
- **Replication check could never fail** - the `\d+\s+fail` regex does not match real
  `repadmin /replsummary` output. Replaced by `Get-ADReplicationFailure`; single-DC
  domains now report INFO.
- **Stale A-record filter matched every record** in the zone as soon as one stale
  name existed (inner `$_` shadowed the record).
- **Autoenrollment GPO false positives** - a text match on `AutoEnrollment` also hit
  GPOs that explicitly *disable* autoenrollment, and unlinked GPOs. Now an XPath check
  on `EnrollCertificatesAutomatically=true`, in an enabled Computer/User half, with at
  least one enabled link.
- **`certutil -viewstore` opens a GUI dialog** (hangs unattended). NTAuth is now read
  from the `NTAuthCertificates` AD object and each cert is checked against dead CAs.
- **CA liveness was English-text dependent** - now judged on certutil's exit code.
- **EVENTS reported PASS for an unreachable DC** (`-ErrorAction SilentlyContinue`
  swallowed the error, count 0). Now only "no matching events" counts as PASS.
- 2889 "no events" detection was English-text dependent; IPv6 client addresses were
  split on the wrong colon.
- Empty `USERDNSDOMAIN` (SYSTEM, scheduled task) left `-Domain` empty - falls back to
  the computer's domain.
- An invalid `Skip` entry in the config aborted config loading - now warned and ignored.

#### Added
- LDAPS: real TLS handshake on 636 with served-cert expiry (a DC listens on 636 even
  with no usable cert); SAN shown when the cert has no Subject.
- 2887 fallback + diagnostics-level readout when per-client 2889 logging is off.
- `DeadCAHostname` config key is now used (reports if it is still registered).
- `-SkipGpoScan` switch for domains with many GPOs.
- Top event IDs in EVENTS detail; sample account names in USERS detail.
- Version stamping (banner, HTML footer, CSV column, returned object).
- `sample.config.json` - anonymous config template with every key.

#### Changed
- Stale users/computers and password-never-expires use server-side LDAP filters
  instead of pulling every account.
- Config key `IverAdminRegex` renamed to `NamedAdminRegex` (default
  `^(adm[-_.]|admin[-_.])`). **Breaking** for configs using the old key.
- All customer-specific names, hosts and defaults removed; author/company lines dropped.
- Functions renamed to approved verbs: `Should-Run` -> `Test-SectionEnabled`,
  `Escape-Html` -> `ConvertTo-HtmlText`.

### [1.0.0] - 2026-09
- Initial consolidated seed (never ran - see 1.1.0).

## Get-CertBlastRadius.ps1

### [1.1.0] - 2026-09-28

#### Fixed
- **Did not parse anywhere.** `$using:probe.ToString()` is not allowed in a using
  expression (PS 7 parse error); non-ASCII without a BOM broke PS 5.1 (5 errors).
- The `-Parallel` block is replaced by a single fan-out `Invoke-Command -ComputerName
  <list> -ThrottleLimit`, which is concurrent on PS 5.1 as well; unreachable hosts are
  reported with the WinRM error.
- The CA's own self-signed certificate is no longer counted as a dependent.
- Inventory `.md` parser no longer picks up OIDs, IPs or version numbers as hosts.

#### Changed
- `HasClientAuth` is also true for Any Purpose and for certs with **no EKU**
  (valid for all purposes). EKUs are read from the raw 2.5.29.37 extension.
- **Breaking:** `-IssuerMatch` is now mandatory - the old default was one specific
  customer's CA.
- `ActiveDirectory` module only required for AD discovery, not with `-ComputerName`.
- `-IssuerMatch` documented as the case-insensitive substring it actually is.
- SAN shown when a cert has no Subject; version stamping; unreachable hosts noted as
  making the result a minimum.

### [1.0.0] - 2026-09
- Initial version (never ran - see 1.1.0).
