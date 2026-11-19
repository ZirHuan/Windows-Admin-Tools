# DomainScanner - Invoke-DomainHealthAudit.ps1 v1.2.0 / Get-CertBlastRadius.ps1 v1.1.0

Ett **read-only** uppsamlingsscript som inventerar och granskar en Windows-domän i ett svep.
Domän-agnostiskt, drivs av parametrar och en valfri JSON-config per miljö. Ingenting ändrar
tillstånd - inget `-Execute`, alltid säkert att köra om.

## Sektioner

| Sektion | Vad den kontrollerar |
|--------|----------------------|
| FOREST | DFL/FFL, UPN-suffix, tombstone lifetime |
| AD | DC-inventering, FSMO-placering, replikeringsfel (`Get-ADReplicationFailure`), SYSVOL, FRS vs DFSR, PDC:ns tidskälla |
| SITES | Sites, subnät, sites utan subnät, DC->site-placering |
| DNS | Namnuppslag av DC, scavenging/aging, stale A-poster |
| CERT | LDAPS-TLS-handskakning + cert-utgång på DC, registrerad CA **+ liveness (certutil -ping, exitkod)**, **länkad autoenrollment-GPO vs död CA**, NTAuth-store (AD-objektet), GPO:er utan länk, valfritt blast radius-svep (`-CertSweep`) |
| USERS | Privilegierade grupper (via SID) + riskklassning, krbtgt-ålder, lösenordspolicy (+ fine-grained), PwdNeverExpires, stale users/computers, AdminSDHolder-föräldralösa, AS-REP roastable, Protected Users |
| SECURITY | Okontrollerad Kerberos-delegering (ej DC), dubbla SPN:er, LAPS-täckning (Windows LAPS / legacy), Kerberoastable konton |
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
`-LogRoot` (`C:\Logs\DomainAudit`), `-SkipGpoScan`, `-CertSweep`.

Alla kontroller är språkoberoende (exitkoder, registervärden, LDAP-filter och SID:er i
stället för tolkad text eller engelska gruppnamn), så de fungerar på svensk Windows.

## Byggstenar

Delar som är användbara på egen hand är separata script som audit-scriptet anropar. De
körs också fristående:

| Script | Används i | Gör |
|--------|-----------|-----|
| `AD\Get-FsmoRoleHolder.ps1` | AD | De fem FSMO-rollerna och vem som har dem |
| `AD\Get-PrivilegedAccountRisk.ps1` | USERS, SECURITY | Alla i privilegierade grupper med OK/REVIEW/HIGH-RISK |
| `AD\Get-UnsignedLdapClient.ps1` | NET | Klienter som gör osignerade/simple LDAP-binds (2889/2887) |
| `Network\Test-TcpPort.ps1` | NET, CERT | TCP-porttest (ersätter Test-OpenPorts.ps1) |
| `Network\Get-TlsCertificate.ps1` | CERT | TLS-handskakning + servercert för valfri host:port |
| `Get-CertBlastRadius.ps1` | CERT (`-CertSweep`) | Maskiner som håller cert från en viss CA |

Scripten letas upp **bredvid audit-scriptet först, sedan i syskonmapparna** (`..\AD`,
`..\Network`). Både repo-strukturen och en platt release-mapp med alla filer fungerar.
Bara byggstenarna för sektionerna som körs krävs; saknas en avbryts körningen direkt med
en lista över vad som saknas (eller `-Skip` de sektionerna).

## Get-CertBlastRadius.ps1

`Get-CertBlastRadius.ps1` - **read-only** svep som hittar varje maskin som håller ett cert
utfärdat av en viss (oftast avvecklad) CA, med utgång, EKU och `HasClientAuth`
(802.1x/EAP-TLS-markören). CERT-sektionen i huvudscriptet flaggar den *döda* CA:n och att
autoenrollment-GPO:er fortfarande pekar på den; det här scriptet mäter *blast radius* -
vilka maskiner som faktiskt beror på den. Se `REVIEW-dead-ca.md` för scenariot.
Med `-CertSweep` kör audit-scriptet det automatiskt för varje död CA.

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
| `DeadCAIssuerMatch` | Issuer-delsträng som `-CertSweep` också sveper efter (utöver CA:er som inte svarar) |
| `ServiceAcctRegex` | Namnmönster för tjänstekonton (HIGH-RISK i högpriv-grupp) |
| `NamedAdminRegex` | Namnmönster för personliga adminkonton (REVIEW i högpriv-grupp). Gamla namnet `AdminAcctRegex` läses också |
| `Skip` | Sektioner att hoppa över; ogiltiga namn varnas och ignoreras |

## Versioner

Båda scripten versioneras var för sig enligt SemVer. Versionen finns i `.NOTES`, i
`$ScriptVersion` (visas i bannern och stämplas i HTML/CSV och returobjektet) och i
[`CHANGELOG.md`](CHANGELOG.md). Git-taggar: `domain-health-audit-vX.Y.Z` och
`cert-blast-radius-vX.Y.Z`. **Varje ändring bumpar versionen på alla tre ställena.**

## Testat (1.2.0, 2026-11-19)

Isolerad labbdomän (Server 2022, en DC + Enterprise CA + RDS-server), Windows PowerShell
5.1.20348 som domänadmin. Parse 0 fel på PS 5.1 och PS 7.4. Testdata lades in så att
varje ny kontroll skulle slå till, och labbet rullades sedan tillbaka till snapshot:

| Testdata | Utfall |
|----------|--------|
| Tjänstekonto med SPN i Domain Admins | Privileged Accounts **FAIL** (HIGH-RISK), Kerberoastable **FAIL** `[PRIVILEGED]` |
| Konto utan Kerberos pre-auth | AS-REP Roastable WARN |
| Konto med adminCount=1 utan skyddad grupp | AdminSDHolder Orphans WARN |
| Namngivet adminkonto i Backup Operators | OK (inte högpriv) - legacy-nyckeln `AdminAcctRegex` lästes |
| Datorkonto med okontrollerad delegering | Unconstrained Delegation WARN (DC:n korrekt undantagen) |
| Olänkad GPO + länkad autoenrollment-GPO | Unlinked GPOs INFO, Autoenrollment vs Dead CA **FAIL** |
| CertSvc stoppad på CA:n | CA Liveness FAIL, NTAuth WARN, `-CertSweep`: 1 host, 4 cert, 3 client-auth, 0 onåbara |
| 2889-loggning nivå 2 + en simple bind | Unsigned LDAP WARN med klientens IP |

Labbets egna verkliga fynd kom också: PDC med Type=NT5DS, inga subnät, scavenging av,
ingen LAPS, lösenordspolicy MinLen 7 / ingen lockout. Dessutom verifierat: repo-layout,
platt release-mapp, `-Skip` utan byggstenar, och tydligt fel när en byggsten saknas.

**Inte verifierat:** flera DC:er/replikering, svensk/lokaliserad Windows (koden är skriven
språkoberoende men inte körd på svensk installation), stora domäner, dubbla SPN:er (AD
blockerar dem vid skrivning sedan 2012 R2, så de gick inte att skapa i labbet).

## Kvar att förfina

1. **Pester-tester** för byggstenarna (de fristående scripten går att testa utan AD).
2. Flera DC:er: replikering och FSMO-placering över flera DC:er är inte körda i labbet.
3. Ingen separat WinRM-preflight; fel rapporteras per DC.
4. Fler kontroller att överväga: tomma OU:n, död CA publicerad som Trusted Root via GPO,
   GPO-länk till borttagen GPO.
