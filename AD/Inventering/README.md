# Invoke-DomainHealthAudit.ps1

Ett **read-only** uppsamlingsscript som inventerar och granskar en Windows-domän i ett svep.
Domän-agnostiskt — drivs av parametrar och en valfri JSON-config per kund/miljö.
Ingenting ändrar tillstånd — inget `-Execute`, alltid säkert att köra om.

## Sektioner

| Sektion | Vad den kontrollerar |
|---------|----------------------|
| FOREST | DFL/FFL, forest-mode, UPN-suffix, tombstone lifetime |
| AD | DC-inventering, FSMO-placering, replikering, SYSVOL, FRS vs DFSR, PDC-tidssynk |
| SITES | Sites, subnät, subnät utan mappning, DC→site-placering |
| DNS | Namnuppslag av DC, scavenging/aging, stale A-poster |
| CERT | LDAPS 636 + cert-utgång på DC, registrerad CA + liveness, autoenrollment-GPO vs död CA, NTAuth-store, olänkade GPO:er |
| USERS | Privilegierade grupper + riskklassning, lösenordspolicy, AdminSDHolder-orphans, AS-REP roastable, krbtgt-ålder, stale users/computers |
| SECURITY | Okontrollerad Kerberos-delegering, dubbla SPN:er, LAPS-närvaro, Kerberoastable konton |
| NET | DC-portar (LDAP/LDAPS/KRB/GC/DNS/SMB), stale-IP på 389, osignerade LDAP-binds (2889) |
| LICENSE | Windows-aktiveringsstatus/KMS på nåbara servrar (best-effort) |
| EVENTS | Kritiska Directory Service- och DFSR-fel senaste N timmar |

## Användning

```powershell
# Full granskning av aktuell domän, endast konsol
.\Invoke-DomainHealthAudit.ps1 -Verbose

# Med kundspecifik miljö-config, HTML + CSV
.\Invoke-DomainHealthAudit.ps1 -Domain customer.local -ConfigPath .\customer.config.json -Html -Csv

# Hoppa över sektioner (t.ex. ingen fjärråtkomst till DC)
.\Invoke-DomainHealthAudit.ps1 -Skip LICENSE,NET
```

Nyckelparametrar: `-EventHoursBack` (24), `-InactiveDays` (90), `-CertExpiryWarnDays` (60),
`-LogRoot` (`C:\Logs\DomainAudit`).

## Följeslagarscript

`Get-CertBlastRadius.ps1` — **read-only** estate-svep som hittar varje maskin som håller
ett cert utfärdat av en viss (oftast avvecklad) CA, med utgång, EKU och `HasClientAuth`
(802.1x/EAP-TLS-markören). CERT-sektionen i huvudscriptet flaggar den *döda* CA:n och att
autoenrollment-GPO:er fortfarande pekar på den; det här scriptet mäter *blast radius* —
vilka maskiner som faktiskt beror på den.

```powershell
.\Get-CertBlastRadius.ps1 -InventoryPath .\01-INVENTORY.md -IssuerMatch 'CN=CA, DC=customer, DC=local'
.\Get-CertBlastRadius.ps1 -IssuerMatch 'ca01' -ComputerName dc01.customer.local,srv01.customer.local
```

## Config per miljö

Kopiera `example.config.json` till `<kund>.config.json` och fyll i kundspecifika fakta
(kända riskkonton, förväntad site, stale-DC/IP, förväntad FSMO-hållare, död CA etc.).
Alla fält är valfria — utan config körs sektionerna i INFO-läge där en förväntan saknas.

## Kvarstående att verifiera mot riktig AD

1. **2889-parsing** — property-index för klient-IP kan variera mellan Server 2016/2019/2022/2025.
2. **certutil -ping-tolkning** — Swedish/localized Windows kan returnera översatta strängar; validera mot faktisk DC.
3. **Remoting-antaganden** — CERT/LICENSE/EVENTS förutsätter WinRM; lägg till pre-flight och snygg degradering.
4. **Prestanda på stora domäner** — `Get-ADUser -Filter {Enabled}` med LastLogonDate drar allt;
   byt till LDAP-filter på `lastLogonTimestamp` för stora kataloger.
5. **Pester-tester** — lägg in mockade AD-cmdlets så sektionerna kan enhetstestas utan DC.
