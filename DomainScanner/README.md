# DomainScanner - Invoke-DomainHealthAudit.ps1 v1.1.0 / Get-CertBlastRadius.ps1 v1.1.0

Ett **read-only** uppsamlingsscript som inventerar och granskar en Windows-domän i ett svep.
Domän-agnostiskt, drivs av parametrar och en valfri JSON-config per miljö. Ingenting ändrar
tillstånd - inget `-Execute`, alltid säkert att köra om.

## Sektioner

| Sektion | Vad den kontrollerar |
|--------|----------------------|
| FOREST | DFL/FFL, UPN-suffix, tombstone lifetime |
| AD | DC-inventering, FSMO-placering, replikeringsfel (`Get-ADReplicationFailure`), SYSVOL, FRS vs DFSR |
| SITES | Sites, subnät, sites utan subnät, DC->site-placering |
| DNS | Namnuppslag av DC, scavenging/aging, stale A-poster |
| CERT | LDAPS-TLS-handskakning + cert-utgång på DC, registrerad CA **+ liveness (certutil -ping, exitkod)**, **länkad autoenrollment-GPO vs död CA**, NTAuth-store (AD-objektet) |
| USERS | Privilegierade grupper + riskklassning, krbtgt-ålder, PwdNeverExpires, stale users/computers |
| NET | DC-portar (LDAP/LDAPS/KRB/GC/DNS/SMB), stale-IP på 389, osignerade LDAP-binds (2889/2887) |
| LICENSE | Windows-aktiveringsstatus/KMS på DC:er (best-effort) |
| EVENTS | Kritiska Directory Service- och DFSR-fel senaste N timmar |

## Användning

```powershell
# Full granskning av aktuell domän, endast konsol
.\Invoke-DomainHealthAudit.ps1 -Verbose

# Med miljö-config (kopia av sample.config.json), HTML + CSV
.\Invoke-DomainHealthAudit.ps1 -Domain contoso.local -ConfigPath .\contoso.config.json -Html -Csv

# Hoppa över sektioner (t.ex. ingen fjärråtkomst till DC)
.\Invoke-DomainHealthAudit.ps1 -Skip LICENSE,NET
```

Nyckelparametrar: `-EventHoursBack` (24), `-InactiveDays` (90), `-CertExpiryWarnDays` (60),
`-LogRoot` (`C:\Logs\DomainAudit`), `-SkipGpoScan`.

## Följeslagarscript

`Get-CertBlastRadius.ps1` - **read-only** svep som hittar varje maskin som håller ett cert
utfärdat av en viss (oftast avvecklad) CA, med utgång, EKU och `HasClientAuth`
(802.1x/EAP-TLS-markören). CERT-sektionen i huvudscriptet flaggar den *döda* CA:n och att
autoenrollment-GPO:er fortfarande pekar på den; det här scriptet mäter *blast radius* -
vilka maskiner som faktiskt beror på den. Se `REVIEW-dead-ca.md` för scenariot.

```powershell
.\Get-CertBlastRadius.ps1 -IssuerMatch 'CN=Contoso-Old-CA' -InventoryPath .\01-INVENTORY.md
.\Get-CertBlastRadius.ps1 -IssuerMatch 'Contoso-Old-CA' -ComputerName srv-app01,srv-db01
```

`-IssuerMatch` är obligatorisk (skiftlägesokänslig delsträng mot certets Issuer).

## Config per miljö

`sample.config.json` är en anonym mall med alla fält. Kopiera den per kund
(`<kund>.config.json`) och **checka inte in kundspecifika config-filer**. Alla fält är
valfria - utan config körs sektionerna i INFO-läge där en förväntan saknas.

| Fält | Betydelse |
|------|-----------|
| `KnownRiskAccounts` | sAMAccountName som alltid klassas HIGH-RISK (leverantör/extern/föräldralös) |
| `ExpectedSites` | Sites som DC:er ska ligga i |
| `StaleDCNames` / `StaleIPs` | Avvecklade DC-namn / IP:er att leta efter i DNS och på port 389 |
| `ExpectedFsmoHolder` | Kortnamn om alla 5 FSMO ska ligga på en DC |
| `DeadCAHostname` | FQDN för en CA som är känd avvecklad (rapporteras om den fortfarande är registrerad) |
| `DeadCAIssuerMatch` | Referens - värdet att ge `Get-CertBlastRadius.ps1 -IssuerMatch` (läses inte av audit-scriptet) |
| `ServiceAcctRegex` | Namnmönster för tjänstekonton (HIGH-RISK i högpriv-grupp) |
| `NamedAdminRegex` | Namnmönster för personliga adminkonton (REVIEW i högpriv-grupp) |
| `Skip` | Sektioner att hoppa över; ogiltiga namn varnas och ignoreras |

## Versioner

Båda scripten versioneras var för sig enligt SemVer. Versionen finns i `.NOTES`, i
`$ScriptVersion` (visas i bannern och stämplas i HTML/CSV och returobjektet) och i
[`CHANGELOG.md`](CHANGELOG.md). Git-taggar: `domain-health-audit-vX.Y.Z` och
`cert-blast-radius-vX.Y.Z`. **Varje ändring bumpar versionen på alla tre ställena.**

## Testat (1.1.0, 2026-09-28)

Körd i en isolerad labbdomän (Server 2022, Windows PowerShell 5.1) både som
SYSTEM och som domänadmin, i friskt läge och med simulerad död CA (CA-tjänsten stoppad,
länkad autoenrollment-GPO, en avstängd och en olänkad GPO som kontroll, en stale A-post).
Alla FAIL/WARN-vägar i CERT slog till korrekt och kontroll-GPO:erna ignorerades. Parse
0 fel på PS 5.1 och PS 7.6.

**Inte verifierat i labbet (en DC, en site):** replikering mellan flera DC:er, 2889 med
loggning påslagen, svensk/lokaliserad Windows, stora domäner.

## Kvar att förfina

1. ~~2889-parsing~~ - IP-delen tolkas nu robust (även IPv6); **fortfarande overifierat
   med riktiga 2889-händelser** (labbet har loggnivå 0).
2. ~~Stale A-record-filtret~~ - omskrivet och verifierat (träffade exakt 1 av N poster).
3. ~~repadmin-parsing~~ - ersatt med `Get-ADReplicationFailure`; flera DC:er ej testat.
4. **Remoting-antaganden** - fel rapporteras nu per DC i stället för falsk PASS, men
   ingen separat WinRM-preflight.
5. ~~Prestanda~~ - LDAP-filter på `lastLogonTimestamp` på serversidan.
6. **Pester-tester** - saknas fortfarande.
7. **Fler kontroller att överväga**: GPO utan länk / länk utan GPO, tomma OU:n,
   `AdminSDHolder`-drift, LAPS-närvaro, tidssynk (w32tm) mot PDC, dubbla SPN:er,
   död CA publicerad som Trusted Root via GPO.
8. `Get-ADGroupMember -Recursive` fallerar på grupper med foreign security principals -
   rapporteras nu som INFO i stället för att tystas.
