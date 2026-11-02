# Dead CA / Autoenrollment: Case Notes Template

**Author:** Christer Rosvall (Iver AB)
**Status:** Discovery complete, remediation **gated on a business decision** (not yet actioned)
**Purpose:** Reference document — anonymized case notes for the dead-CA + autoenrollment scenario.

---

## 1. What was found (confirmed facts)

A report claimed "a GPO points CA autoenrollment at an old CA box that no longer exists."
Investigation confirmed the shape of it, with corrections:

- **One CA is registered in AD** — `Name=CA`, `dNSHostName=ca01.CUSTOMER.local` — in
  `CN=Enrollment Services,CN=Public Key Services,CN=Services,<configNC>`. Published cert name **AD-CA01-CA**.
- **That CA does not answer.** `certutil -ping` fails forest-wide
  (`0x80070103 ERROR_NO_MORE_ITEMS`, "No local Certification Authority").
- **No replacement CA exists.** ADCS is *Available* (not Installed) on `<DC01>`;
  `CertSvc` absent. `<DC01>` only holds the LDAPS cert — it is **not** a CA.
- **Two autoenrollment GPOs, both linked at the domain root** (SOMPath = `CUSTOMER.local`, no OU),
  security filtering = Authenticated Users → **whole estate in scope**:
  - `CA Auto Enrollment`
  - `<prefix> 802.1x CA autoenrollment`
  - `CA Auto Enrollment` also publishes **AD-CA01-CA as a Trusted Root** (expiry: check GPO).
- **NTAuth store is empty** (`certutil -viewstore -enterprise NTAuth` returns success, no certs)
  — the dead CA is *not* sitting as a trusted enterprise issuer. Less stale trust to clean than feared.
- **Dependency is NOT uniform across the estate:**
  - `<DC01>` holds several certs issued by `CN=CA, DC=CUSTOMER, DC=local`, **expiring `<DATE>`**.
  - `<DC02>` holds **none** from that issuer — only device-registration / subjectless certs.
- **Urgency:** slow burn. The load-bearing machine certs expire `<DATE>`. They will **not renew** —
  no CA to renew against. Not a fire; a fuse.

### The open decision (gates all remediation)

Two questions for the customer, still unanswered:
1. Is wired/wireless **802.1x certificate-based (EAP-TLS)?** The `802.1x` GPO implies yes, but
   implication ≠ confirmation.
2. Does the customer **want a CA going forward**, or should cert-based auth be retired?

- If **certs needed** → rebuild-CA project (stand up ADCS, migrate/republish templates,
  re-point autoenrollment, reissue before `<DATE>`), then retire old CA references.
- If **certs not needed** → disable both GPOs, remove the stale `CN=CA` Enrollment Services object,
  purge the dead root from the Trusted-Root GPO + any CDP/AIA leftovers.
  **Destructive — only after confirming nothing depends on the dead CA.**

> Nothing destructive has been run. Log this as a customer action-list item.

---

## 2. Script coverage of this scenario

### `Invoke-DomainHealthAudit.ps1` — CERT section

- **Enterprise CA Registered** — lists each `pKIEnrollmentService` with its `dNSHostName`.
- **CA Liveness `<name>`** — TCP/135 probe + `certutil -ping -config <host>\<caName>`.
  Emits **FAIL** when a registered CA does not answer.
- **Autoenrollment vs Dead CA** — enumerates all GPOs, flags autoenrollment markers,
  cross-references against dead CAs. **FAIL** when autoenrollment GPOs are active and a
  registered CA is dead.
- **NTAuth Store** — counts certs in the enterprise NTAuth store.
- **Unlinked GPOs** — flags GPOs with no OU/site/domain link.

### `Get-CertBlastRadius.ps1` — estate cert sweep

- Targets from `-ComputerName`, an inventory file, or all enabled Windows computers in AD.
- Remotes to each host, reads `Cert:\LocalMachine\My`, keeps certs whose **Issuer** matches
  `-IssuerMatch`.
- Per cert: Subject, Issuer, NotAfter, DaysToExpiry, EKU, **HasClientAuth** (802.1x marker), Thumbprint.
- Loud **red warning** if any Client-Auth certs are found: do not retire the CA before handling them.

---

## 3. Known validation gaps (verify against real AD before relying on output)

1. **`certutil -ping` parsing** — matches `interface is alive` / `completed successfully` as strings.
   Localized (non-English) Windows may return translated output — validate, or rely on `$LASTEXITCODE`.
2. **GPO autoenrollment detection** is XML string-matching (`AutoEnrollment` / `Auto-Enrollment`).
   Confirm markers appear in `Get-GPOReport -ReportType Xml` for a real autoenrollment policy;
   XPath against the Public Key policy node would be sturdier.
3. **`Get-GPO -All` + per-GPO XML** is heavy on large domains — skip with `-Skip CERT` if needed.
4. **NTAuth counting** via `Serial Number:` regex — verify count is correct (and 0 for empty store)
   across locales.
5. **PS7 parallel block** in `Get-CertBlastRadius.ps1` — `[scriptblock]::Create($using:probe.ToString())`
   should work but verify; fall back to sequential if unstable.
6. **Remoting assumptions** — CERT liveness uses `certutil` locally against remote configs; the sweep
   uses `Invoke-Command`. Add WinRM pre-flight and graceful degrade where blocked.
7. **HasClientAuth OID** — `EnhancedKeyUsageList.ObjectId -contains '1.3.6.1.5.5.7.3.2'` may need the
   cert re-read with the EKU extension on PS5; a cert with *no* EKU (all-purpose) should arguably also
   be flagged.

---

## 4. Do NOT let the scripts drift into

- **Anything destructive.** Both scripts are read-only by design.
- **Auto-deciding retire-vs-rebuild.** The scripts *report*; the decision is the customer's.
