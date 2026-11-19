# Dead CA + Live Autoenrollment - Scenario and Review Notes

**Status:** Reference scenario for `Invoke-DomainHealthAudit.ps1` (CERT section) and
`Get-CertBlastRadius.ps1`. All names below are placeholders (`contoso.local`).

---

## 1. The scenario

A typical finding: "a GPO points certificate autoenrollment at an old CA server that no
longer exists." When investigated, the shape is usually:

- **One enterprise CA is still registered in AD** - a `pKIEnrollmentService` object in
  `CN=Enrollment Services,CN=Public Key Services,CN=Services,<configNC>`, e.g.
  `Name=Contoso-Old-CA`, `dNSHostName=oldca.contoso.local`.
- **That CA does not answer.** `certutil -ping -config "oldca.contoso.local\Contoso-Old-CA"`
  fails everywhere (e.g. `0x800706BA RPC_S_SERVER_UNAVAILABLE`).
- **No replacement CA exists** - AD CS is not installed anywhere else.
- **Autoenrollment GPOs are still linked at the domain root** with security filtering
  Authenticated Users, so the **whole estate is in scope**. The same GPO often also
  publishes the dead CA's certificate as a **Trusted Root**.
- **NTAuth** may or may not still contain the dead CA's certificate.
- **Dependency is not uniform:** some machines hold certs issued by the dead CA (often
  with Client Authentication EKU = 802.1x / EAP-TLS), others hold none.
- **Urgency is a fuse, not a fire:** issued certs keep working until they expire, then
  they cannot renew because there is no CA to renew against.

### The decision that gates remediation (belongs to the customer)
1. Is wired/wireless **802.1x certificate-based (EAP-TLS)**?
2. Does the customer **want a CA going forward**, or should cert-based auth be retired?

- **Certs needed** -> rebuild-CA project (install AD CS, publish templates, re-point
  autoenrollment, reissue before the old certs expire), then retire the old references.
- **Certs not needed** -> disable the autoenrollment GPOs, remove the stale
  Enrollment Services object, purge the dead root from the Trusted-Root GPO and any
  CDP/AIA leftovers. **Destructive - only after confirming nothing depends on the CA.**

---

## 2. What the tools do

### `Invoke-DomainHealthAudit.ps1` - CERT section
- **Enterprise CA Registered** - lists each `pKIEnrollmentService` with its host.
- **CA Liveness** - TCP/135 probe plus `certutil -ping`, judged on the **exit code**
  (locale-independent). FAIL when a registered CA does not answer.
- **Autoenrollment vs Dead CA** - XPath on each GPO report: counts a GPO only when
  `EnrollCertificatesAutomatically=true`, in an enabled Computer/User half, with at
  least one enabled link. FAIL when such a GPO exists and a registered CA is dead.
- **NTAuth Store** - read from the `NTAuthCertificates` AD object; WARN if a cert
  belongs to a dead CA.
- **Known Dead CA Registered** - when `DeadCAHostname` is set in the config.

### `Get-CertBlastRadius.ps1` - dependency sweep
- Targets from `-ComputerName`, an inventory file, or all enabled Windows computers in AD.
- Reads `Cert:\LocalMachine\My` remotely and keeps certs whose Issuer contains
  `-IssuerMatch` (the CA's own self-signed cert is excluded).
- Per cert: Subject (or SAN), Issuer, NotAfter, DaysToExpiry, EKU, **HasClientAuth**
  (Client Auth, Any Purpose, or no EKU at all), Thumbprint. CSV output.
- Red warning when client-auth-capable certs depend on the CA; unreachable hosts are
  listed and make the result a minimum.

---

## 3. Review outcome (v1.1.0, 2026-09-28)

Verified in an isolated Server 2022 lab (Windows PowerShell 5.1) with a simulated dead
CA (CA service stopped) plus ON / OFF / unlinked autoenrollment test GPOs. Full detail in
`CHANGELOG.md`.

1. **certutil -ping** - exit code 0 when alive; a dead CA returned `-2147023174`
   (0x800706BA). No text matching.
2. **GPO detection** - a text match on `AutoEnrollment` was wrong: a GPO that
   *disables* autoenrollment contains `AutoEnrollmentSettings` too. The XPath check
   counted only the ON GPO.
3. **Heavy GPO scan** - `-SkipGpoScan`; a missing GroupPolicy module degrades to INFO.
4. **NTAuth** - `certutil -viewstore` opens a GUI dialog; replaced by the AD object.
5. **PS7 parallel block** - was a parse error; replaced by fan-out `Invoke-Command`.
6. **Remoting** - unreachable hosts/DCs are reported, not counted as clean/PASS.
7. **HasClientAuth** - EKUs read from the raw extension; no-EKU and Any Purpose certs
   count as client-auth capable.
8. **Syntax** - both scripts now parse with 0 errors on PS 5.1 and 7.
   PSScriptAnalyzer has **not** been run yet.

Not verified in the lab: multiple DCs / replication, real 2889 events, localized
(non-English) Windows, large domains.

---

## 4. Do NOT let the scripts drift into

- **Anything destructive.** Both are read-only by design. GPO disable, stale-object
  removal and NTAuth purge are a separate, gated deliverable.
- **Auto-deciding retire-vs-rebuild.** The scripts report; the customer decides.
