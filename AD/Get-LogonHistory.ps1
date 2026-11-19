#Requires -Version 5.1
<#
.SYNOPSIS
    Rapport over vilka som har loggat in pa en Windows-dator: antal, forsta och senaste inloggning per anvandare.

.DESCRIPTION
    Laser handelseloggarna pa den lokala datorn (eller en annan dator med -ComputerName) och
    sammanstaller inloggningar per anvandare. Fungerar pa Windows-servrar, Windows-klienter
    och datorer utan domain (workgroup), eftersom grundkallan ar Security-loggen.

    Du valjer vilka delar som ska sokas igenom med -Source:

      Logon        Security-loggen, handelse 4624 (lyckad inloggning). Fungerar pa alla
                   Windows-datorer. Vilka inloggningstyper som raknas styrs av -LogonType.
                   Standardkalla.
      RdsSession   TerminalServices-LocalSessionManager, handelse 21 (lyckad RDS/RDP-
                   sessionsinloggning). Med -IncludeReconnect raknas aven handelse 25
                   (aterkoppling till befintlig session).
      RdpAuth      TerminalServices-RemoteConnectionManager, handelse 1149 (RDP-
                   autentisering lyckades). Visar kallans IP-adress.
      FailedLogon  Security-loggen, handelse 4625 (misslyckad inloggning).
      NpsAccess    Security-loggen, handelse 6272 (Network Policy Server beviljade atkomst).
                   Pa en server med RD Gateway ar det en post per uppkoppling via gatewayen,
                   med kontonamn och klientdatorns namn. Skrivs oberoende av "Audit Logon".
      All          Alla ovan.

    Rapporten visar forst hur langt bak i tiden varje logg gar (aldsta handelse, storlek,
    overskrivningslage), sedan en tabell per kalla och anvandare. Resultatet sparas som
    semikolonseparerad CSV (oppnas direkt i Excel) om du inte anger -NoCsv.

    Systemkonton (maskinkonton som slutar pa $, SYSTEM, LOCAL SERVICE, NETWORK SERVICE,
    DWM-/UMFD-konton, ANONYMOUS LOGON) tas bort ur Security-kallorna om du inte anger
    -IncludeSystemAccounts.

.PARAMETER Source
    Vilka delar som ska sokas: Logon, RdsSession, RdpAuth, FailedLogon, NpsAccess eller All.
    Flera kan anges separerade med komma. Standard: Logon.

.PARAMETER LogonType
    Vilka inloggningstyper som raknas for kallan Logon (handelse 4624):
      Interactive (2)         konsol / lokal inloggning
      RemoteInteractive (10)  RDP / Remote Desktop
      CachedInteractive (11)  inloggning med cachade uppgifter (dator utan kontakt med domankontrollant)
      Unlock (7)              upplasning av last session (RDP-aterkoppling loggas ofta som denna typ)
      Network (3)             natverksinloggning (fildelning m.m. - mycket brus)
      Batch (4)               schemalagda uppgifter
      Service (5)             tjanster
      NetworkCleartext (8)    natverksinloggning med klartextuppgifter (t.ex. IIS Basic)
      NewCredentials (9)      RunAs /netonly
    Standard: Interactive, RemoteInteractive, CachedInteractive.
    For FailedLogon (4625) anvands inget typfilter om du inte anger -LogonType uttryckligen.

.PARAMETER Since
    Tidigaste tidpunkt som ska med. Utan -Since och -Days tas allt som finns kvar i loggen med.

.PARAMETER Until
    Senaste tidpunkt som ska med. Standard: nu.

.PARAMETER Days
    Hamta de senaste N dagarna. Kan inte kombineras med -Since.

.PARAMETER IncludeReconnect
    Raknar aven handelse 25 (aterkoppling) for kallan RdsSession.

.PARAMETER User
    Visa bara anvandare som matchar nagot av monstren (wildcard, t.ex. 'CONTOSO\a-*' eller 'anna*').
    Matchas mot bade DOMAN\anvandare och enbart anvandarnamnet.

.PARAMETER ExcludeUser
    Dolj anvandare som matchar nagot av monstren (wildcard, t.ex. 'svc-*').

.PARAMETER IncludeSystemAccounts
    Ta med systemkonton i Security-kallorna (Logon, FailedLogon).

.PARAMETER Detailed
    Visa varje enskild handelse (tid, anvandare, IP, detalj) i stallet for sammanstallning per anvandare.

.PARAMETER PassThru
    Skicka resultatobjekten vidare i pipelinen i stallet for att skriva ut en tabell.
    Anvands for egen vidarebearbetning, t.ex. | Where-Object Antal -gt 10.

.PARAMETER ComputerName
    Dator att lasa loggarna fran. Utelamnas parametern kors skriptet mot den lokala datorn.
    Kraver att "Remote Event Log Management" ar tillatet i brandvaggen pa maldatorn.

.PARAMETER Credential
    Uppgifter som anvands mot -ComputerName.

.PARAMETER SkipLogInfo
    Hoppa over avsnittet om loggdjup (aldsta handelse, storlek m.m.).

.PARAMETER OutCsv
    Sokvag till CSV-filen. Standard: LogonReport_<dator>_<datum_tid>.csv i aktuell mapp.

.PARAMETER NoCsv
    Spara ingen CSV-fil.

.EXAMPLE
    .\Get-LogonHistory.ps1

    Standardkorning pa den lokala datorn: konsol-, RDP- och cachade inloggningar per anvandare
    fran Security-loggen, hela loggen bakat, sparas som CSV i aktuell mapp.

.EXAMPLE
    .\Get-LogonHistory.ps1 -Source RdsSession -IncludeReconnect

    Pa en RDS-server: RDS-sessionsinloggningar inklusive aterkopplingar, per anvandare.

.EXAMPLE
    .\Get-LogonHistory.ps1 -Source RdsSession,RdpAuth,Logon -Days 30

    Jamfor tre kallor for RDP under de senaste 30 dagarna. RdpAuth visar kallans IP-adress.

.EXAMPLE
    .\Get-LogonHistory.ps1 -Source NpsAccess,RdsSession -Days 30

    Pa en server med RD Gateway: uppkopplingar via gatewayen (konto + klientdatorns namn)
    jamford med RDS-sessionsinloggningar, senaste 30 dagarna. Anvandbart nar "Audit Logon"
    ar avstangt och Security 4624 saknas. Lagg till -Verbose for att se faltnamnen i 6272.

.EXAMPLE
    .\Get-LogonHistory.ps1 -Source All -Days 7 -NoCsv

    Alla kallor (inklusive misslyckade inloggningar) senaste veckan, ingen fil sparas.

.EXAMPLE
    .\Get-LogonHistory.ps1 -Source FailedLogon -Days 1 -Detailed

    Varje misslyckat inloggningsforsok senaste dygnet med tid, konto, kalladress och orsakskod.

.EXAMPLE
    .\Get-LogonHistory.ps1 -LogonType RemoteInteractive,Unlock -User 'CONTOSO\a-*' -ExcludeUser '*test*'

    Bara RDP-inloggningar och -upplasningar for konton som borjar pa a- i doman CONTOSO,
    utan konton som innehaller "test".

.EXAMPLE
    .\Get-LogonHistory.ps1 -Since '2026-10-01' -Until '2026-10-31 23:59'

    Inloggningar under oktober 2026.

.EXAMPLE
    .\Get-LogonHistory.ps1 -ComputerName SRV01 -Credential (Get-Credential) -Source RdsSession

    Las RDS-sessionsloggen fran en annan dator. Om den inte kan nas pa distans: kor skriptet lokalt.

.EXAMPLE
    .\Get-LogonHistory.ps1 -PassThru -NoCsv | Where-Object Antal -gt 10 | Sort-Object Antal -Descending

    Skicka objekten vidare och visa bara anvandare med fler an 10 inloggningar.

.NOTES
    Version : 1.2.0
    Kraver  : Windows PowerShell 5.1 eller PowerShell 7. Kor som administrator (lokal admin) -
              Security-loggen kan inte lasas utan det. Skriptet varnar om rattigheter saknas.
              Skriptet andrar ingenting pa datorn, det bara laser handelseloggar.

    Changelog
      1.2.0  2026-11-19  Kontrollerar "Audit Logon" (auditpol) nar Logon/FailedLogon soks pa den
                         lokala datorn. Ar den avstangd visas "Audit Logon ar avstangd", hur man
                         slar pa den (GPO / auditpol + lank) och alternativet NpsAccess,
                         RdsSession, RdpAuth. Samma tips visas om Logon/FailedLogon gav 0 traffar.
      1.1.0  2026-11-06  Ny kalla NpsAccess (Security 6272, RD Gateway / NPS). Fungerar aven
                         nar "Audit Logon" ar avstangt.
      1.0.0  2026-11-06  Forsta versionen. Generaliserad fran ett kundspecifikt skript som bara
                         lastes RDS-sessionsloggen pa en namngiven server. Nu: lokal dator som
                         standard, valbara kallor (-Source), Security-handelse 4624/4625,
                         RDP-autentisering (1149), filter pa tid och anvandare, detaljlage.

    Att tanka pa
      * Hur langt bak man kan se bestams av loggens maxstorlek och overskrivningslage.
        Cirkulara loggar skriver over de aldsta handelserna nar de ar fulla. Avsnittet
        "Loggdjup" visar hur langt bak varje logg gar - resultatet galler bara den perioden.
      * Security 4624/4625 forutsatter att granskning ("Audit Logon") ar paslagen for Success
        respektive Failure. Skriptet kontrollerar detta sjalvt pa den lokala datorn. Manuellt
        (fungerar oavsett sprak pa Windows):
          auditpol /get /subcategory:"{0CCE9215-69AE-11D9-BED3-505054503030}"
      * Kalladressen kan saknas eller vara ett platshallarvarde. Pa en RDS-server med
        RD Gateway visade handelse 21 "::%16777216" i stallet for klientens IP, medan
        6272 (NpsAccess) inneholl klientdatorns namn men ingen IP. Skriptet visar
        vardet som det star i loggen. Klientens IP kan finnas i gatewayens egen logg
        (TerminalServices-Gateway/Operational) - inte inlagd i detta skript.
      * NpsAccess raknar uppkopplingar via gatewayen, inte sessioner: en och samma session
        kan ge flera poster. Jamfor med RdsSession.
      * En RDP-inloggning kan synas i flera kallor samtidigt (RdsSession, RdpAuth och Logon
        typ 10). Kallorna summeras inte ihop - de visas som separata rader.
      * Uppkoppling mot en dator utan domain (workgroup) pa distans fungerar ofta daligt
        (lokal kontofiltrering). Kor da skriptet lokalt pa datorn.
      * Tider ar datorns lokala tid.
#>
[CmdletBinding()]
param(
    [ValidateSet('Logon', 'RdsSession', 'RdpAuth', 'FailedLogon', 'NpsAccess', 'All')]
    [string[]]$Source = @('Logon'),

    [ValidateSet('Interactive', 'Unlock', 'RemoteInteractive', 'CachedInteractive', 'Network', 'Batch', 'Service', 'NetworkCleartext', 'NewCredentials')]
    [string[]]$LogonType = @('Interactive', 'RemoteInteractive', 'CachedInteractive'),

    [datetime]$Since,
    [datetime]$Until,

    [ValidateRange(1, 36500)]
    [int]$Days,

    [switch]$IncludeReconnect,
    [string[]]$User,
    [string[]]$ExcludeUser,
    [switch]$IncludeSystemAccounts,
    [switch]$Detailed,
    [switch]$PassThru,

    [string]$ComputerName,
    [System.Management.Automation.PSCredential]$Credential,

    [switch]$SkipLogInfo,
    [string]$OutCsv,
    [switch]$NoCsv
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Tabeller
# ---------------------------------------------------------------------------
$logonTypeId = @{
    Interactive = 2; Network = 3; Batch = 4; Service = 5; Unlock = 7
    NetworkCleartext = 8; NewCredentials = 9; RemoteInteractive = 10; CachedInteractive = 11
}
$logonTypeName = @{
    2 = 'Interactive'; 3 = 'Network'; 4 = 'Batch'; 5 = 'Service'; 7 = 'Unlock'
    8 = 'NetworkCleartext'; 9 = 'NewCredentials'; 10 = 'RemoteInteractive'; 11 = 'CachedInteractive'
    12 = 'CachedRemoteInteractive'; 13 = 'CachedUnlock'
}
$logNames = @{
    Logon       = 'Security'
    FailedLogon = 'Security'
    NpsAccess   = 'Security'
    RdsSession  = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'
    RdpAuth     = 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational'
}

# ---------------------------------------------------------------------------
# Indata: kallor, tidsfonster, mal
# ---------------------------------------------------------------------------
if ($Source -contains 'All') {
    $sources = @('Logon', 'RdsSession', 'RdpAuth', 'FailedLogon', 'NpsAccess')
}
else {
    $sources = @($Source | Select-Object -Unique)
}

$from = $null
$to = $null
if ($PSBoundParameters.ContainsKey('Since')) { $from = $Since }
if ($PSBoundParameters.ContainsKey('Days')) {
    if ($null -ne $from) { throw 'Anvand antingen -Days eller -Since, inte bada.' }
    $from = (Get-Date).AddDays(-$Days)
}
if ($PSBoundParameters.ContainsKey('Until')) { $to = $Until }
if ($null -ne $from -and $null -ne $to -and $from -ge $to) { throw '-Since/-Days maste vara tidigare an -Until.' }

$remote = @{}
if ($ComputerName) { $remote['ComputerName'] = $ComputerName }
if ($Credential) {
    if (-not $ComputerName) { throw '-Credential kraver -ComputerName.' }
    $remote['Credential'] = $Credential
}
$target = if ($ComputerName) { $ComputerName } else { $env:COMPUTERNAME }

$typeFilterGiven = $PSBoundParameters.ContainsKey('LogonType')
$typeIds = @($LogonType | ForEach-Object { $logonTypeId[$_] })

# ---------------------------------------------------------------------------
# Hjalpfunktioner
# ---------------------------------------------------------------------------
function Test-IsAdministrator {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function New-EventXPath {
    param([int[]]$EventId, [int[]]$TypeIds)

    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $idExpr = ($EventId | ForEach-Object { "EventID=$_" }) -join ' or '
    $sys = "($idExpr)"
    if ($null -ne $from) {
        $sys += " and TimeCreated[@SystemTime>='" + $from.ToUniversalTime().ToString('yyyy-MM-dd\THH:mm:ss.fff\Z', $inv) + "']"
    }
    if ($null -ne $to) {
        $sys += " and TimeCreated[@SystemTime<='" + $to.ToUniversalTime().ToString('yyyy-MM-dd\THH:mm:ss.fff\Z', $inv) + "']"
    }
    $xp = "*[System[$sys]]"
    if ($TypeIds -and $TypeIds.Count -gt 0) {
        $t = ($TypeIds | ForEach-Object { "Data[@Name='LogonType']='$_'" }) -join ' or '
        $xp += " and *[EventData[($t)]]"
    }
    return $xp
}

function Invoke-EventQuery {
    param([string]$LogName, [string]$XPath)

    $p = @{ LogName = $LogName; FilterXPath = $XPath; ErrorAction = 'Stop' }
    foreach ($k in $remote.Keys) { $p[$k] = $remote[$k] }
    try {
        Get-WinEvent @p
    }
    catch {
        # Inga traffar ar inget fel - returnera tomt. Allt annat kastas vidare.
        if ($_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*') { return @() }
        throw
    }
}

function Test-SystemAccount {
    param([string]$Name, [string]$Domain, [string]$Sid)

    if ([string]::IsNullOrWhiteSpace($Name) -or $Name -eq '-') { return $true }
    if ($Name.EndsWith('$')) { return $true }
    if ($Name -like 'DWM-*' -or $Name -like 'UMFD-*') { return $true }
    if ($Name -in @('SYSTEM', 'LOCAL SERVICE', 'NETWORK SERVICE', 'ANONYMOUS LOGON')) { return $true }
    if ($Domain -in @('NT AUTHORITY', 'Window Manager', 'Font Driver Host', 'NT SERVICE')) { return $true }
    # Sprakoberoende kontroll pa SID: SYSTEM, LOCAL SERVICE, NETWORK SERVICE, anonym, DWM, UMFD
    if ($Sid -and ($Sid -in @('S-1-5-18', 'S-1-5-19', 'S-1-5-20', 'S-1-5-7') -or $Sid -like 'S-1-5-90-*' -or $Sid -like 'S-1-5-96-*')) { return $true }
    return $false
}

function Join-AccountName {
    param([string]$Domain, [string]$Name)
    if ([string]::IsNullOrWhiteSpace($Domain) -or $Domain -eq '-') { return $Name }
    return "$Domain\$Name"
}

function Test-UserWanted {
    param([string]$UserName)

    $short = $UserName.Split('\')[-1]
    if ($User) {
        $ok = $false
        foreach ($pattern in $User) {
            if ($UserName -like $pattern -or $short -like $pattern) { $ok = $true; break }
        }
        if (-not $ok) { return $false }
    }
    foreach ($pattern in $ExcludeUser) {
        if ($UserName -like $pattern -or $short -like $pattern) { return $false }
    }
    return $true
}

function New-Rec {
    param($Ev, [string]$Src, [string]$UserName, [string]$Ip, [string]$Detail)
    [pscustomobject]@{
        Time     = $Ev.TimeCreated
        Computer = $Ev.MachineName
        Source   = $Src
        User     = $UserName
        IP       = $Ip
        Detail   = $Detail
    }
}

function ConvertTo-Record {
    param($Ev, [string]$Src)

    $xml = [xml]$Ev.ToXml()

    if ($Src -eq 'RdsSession') {
        $x = $xml.Event.UserData.EventXML
        $kind = if ($Ev.Id -eq 25) { 'Aterkoppling' } else { 'Inloggning' }
        return (New-Rec $Ev $Src ([string]$x.User) ([string]$x.Address) ("$kind, session $($x.SessionID)"))
    }

    if ($Src -eq 'RdpAuth') {
        # 1149: Param1 = anvandare, Param2 = doman, Param3 = kalladress
        $x = $xml.Event.UserData.EventXML
        $name = Join-AccountName ([string]$x.Param2) ([string]$x.Param1)
        return (New-Rec $Ev $Src $name ([string]$x.Param3) 'RDP-autentisering')
    }

    if ($Src -eq 'NpsAccess') {
        # 6272: konto i Subject*, klientdator i SubjectMachineName, ev. IP i CallingStationID
        $nmap = @{}
        foreach ($d in $xml.Event.EventData.Data) {
            if ($d.Name) { $nmap[[string]$d.Name] = [string]$d.'#text' }
        }
        if (-not $script:npsFieldsShown) {
            Write-Verbose ('NpsAccess falt: ' + (($nmap.Keys | Sort-Object) -join ', '))
            $script:npsFieldsShown = $true
        }
        $nm = $nmap['FullyQualifiedSubjectUserName']
        if ([string]::IsNullOrWhiteSpace($nm) -or $nm -eq '-') {
            $nm = $nmap['SubjectUserName']
            if ($nm -and $nm -notlike '*\*') { $nm = Join-AccountName $nmap['SubjectDomainName'] $nm }
        }
        if ([string]::IsNullOrWhiteSpace($nm) -or $nm -eq '-') { return $null }
        $machine = $nmap['SubjectMachineName']
        if ([string]::IsNullOrWhiteSpace($machine) -or $machine -eq '-') { $machine = $nmap['FullyQualifiedSubjectMachineName'] }
        $nip = [string]$nmap['CallingStationID']
        if ($nip -eq '-') { $nip = '' }
        return (New-Rec $Ev $Src $nm $nip ("Klientdator: $machine, policy: $($nmap['NetworkPolicyName'])"))
    }

    # Logon (4624) och FailedLogon (4625)
    $map = @{}
    foreach ($d in $xml.Event.EventData.Data) {
        if ($d.Name) { $map[[string]$d.Name] = [string]$d.'#text' }
    }
    $uname = $map['TargetUserName']
    $udom = $map['TargetDomainName']
    if (-not $IncludeSystemAccounts) {
        if (Test-SystemAccount -Name $uname -Domain $udom -Sid $map['TargetUserSid']) { return $null }
    }
    elseif ([string]::IsNullOrWhiteSpace($uname) -or $uname -eq '-') {
        return $null
    }

    $lt = 0
    [void][int]::TryParse([string]$map['LogonType'], [ref]$lt)
    $ltText = if ($logonTypeName.ContainsKey($lt)) { "$($logonTypeName[$lt]) ($lt)" } else { "Typ $lt" }
    $ip = [string]$map['IpAddress']
    if ($ip -eq '-') { $ip = '' }

    if ($Src -eq 'FailedLogon') {
        $detail = "$ltText, Status $($map['Status']), SubStatus $($map['SubStatus'])"
    }
    else {
        $detail = $ltText
    }
    return (New-Rec $Ev $Src (Join-AccountName $udom $uname) $ip $detail)
}

$auditLogonGuid = '{0CCE9215-69AE-11D9-BED3-505054503030}'
$auditDocUrl = 'https://learn.microsoft.com/windows-server/identity/ad-ds/plan/security-best-practices/advanced-audit-policy-configuration'
$scriptLeaf = if ($PSCommandPath) { Split-Path -Leaf $PSCommandPath } else { 'Get-LogonHistory.ps1' }

function Get-AuditLogonSetting {
    # Effektiv installning for underkategorin Logon pa den lokala datorn:
    # 0 = av, 1 = Success, 2 = Failure, 3 = bada. $null = kunde inte lasas.
    # Laser kolumnen "Setting Value" ur auditpol /backup - ett tal, oberoende av Windows-sprak.
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('auditpol_{0}.csv' -f [guid]::NewGuid())
    try {
        $null = & auditpol.exe /backup "/file:$tmp" 2>&1
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $tmp)) { return $null }
        # Kolumner: Machine Name, Policy Target, Subcategory, Subcategory GUID,
        #           Inclusion Setting, Exclusion Setting, Setting Value
        $rows = Import-Csv -LiteralPath $tmp -Header 'c0', 'c1', 'c2', 'c3', 'c4', 'c5', 'c6'
        $row = $rows | Where-Object { $_.c3 -eq $auditLogonGuid } | Select-Object -First 1
        if (-not $row) { return $null }
        $v = 0
        if ([int]::TryParse(([string]$row.c6).Trim(), [ref]$v)) { return $v }
        return $null
    }
    catch {
        Write-Verbose "Kunde inte lasa auditpol: $($_.Exception.Message)"
        return $null
    }
    finally {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
}

function Show-AuditLogonHelp {
    # Instruktion: sla pa Audit Logon, eller anvand kallor som inte beror pa den.
    Write-Host '  Sa slar du pa granskningen (galler bara framat - det som inte loggats gar inte att fa fram):' -ForegroundColor Yellow
    Write-Host '    GPO (rekommenderat pa en domanmedlem):'
    Write-Host '      Computer Configuration > Policies > Windows Settings > Security Settings >'
    Write-Host '      Advanced Audit Policy Configuration > Audit Policies > Logon/Logoff >'
    Write-Host '      Audit Logon = Success, Failure'
    Write-Host '      Security Options > "Audit: Force audit policy subcategory settings ..." = Enabled'
    Write-Host '    Lokalt (snabbtest - kan skrivas over av GPO vid nasta gpupdate):'
    Write-Host ('      auditpol /set /subcategory:"{0}" /success:enable /failure:enable' -f $auditLogonGuid)
    Write-Host ('    Mer info: {0}' -f $auditDocUrl)
    Write-Host '  Alternativ nu - kallor som inte beror pa Audit Logon:' -ForegroundColor Yellow
    Write-Host ('    .\{0} -Source NpsAccess,RdsSession,RdpAuth -IncludeReconnect' -f $scriptLeaf)
    Write-Host '      NpsAccess  = uppkopplingar via RD Gateway / NPS (Security 6272, konto + klientdator)'
    Write-Host '      RdsSession = RDS-sessionsinloggningar, RdpAuth = RDP-autentisering med kalladress'
}

# ---------------------------------------------------------------------------
# Rubrik
# ---------------------------------------------------------------------------
Write-Host ''
Write-Host "=== Inloggningsrapport: $target ===" -ForegroundColor Cyan
if (-not $ComputerName) {
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $role = switch ([int]$cs.DomainRole) {
            0 { 'Fristaende klient' }
            1 { 'Klient i doman' }
            2 { 'Fristaende server' }
            3 { 'Medlemsserver' }
            4 { 'Domankontrollant (backup)' }
            5 { 'Domankontrollant (primar)' }
            default { 'Okand roll' }
        }
        $member = if ($cs.PartOfDomain) { "doman $($cs.Domain)" } else { "workgroup $($cs.Domain)" }
        Write-Host ("  {0} | {1} | {2}" -f $os.Caption.Trim(), $role, $member)
    }
    catch {
        Write-Verbose "Kunde inte lasa datorinfo: $($_.Exception.Message)"
    }
    if (($sources -contains 'Logon' -or $sources -contains 'FailedLogon' -or $sources -contains 'NpsAccess') -and -not (Test-IsAdministrator)) {
        Write-Warning 'Skriptet kors inte som administrator - Security-loggen kan inte lasas. Starta PowerShell som administrator.'
    }
}
Write-Host ("  Kallor : {0}" -f ($sources -join ', '))
$periodText = if ($null -ne $from) { $from.ToString('yyyy-MM-dd HH:mm') } else { 'aldsta handelse' }
$periodText += ' - '
$periodText += if ($null -ne $to) { $to.ToString('yyyy-MM-dd HH:mm') } else { 'nu' }
Write-Host ("  Period : {0}" -f $periodText)

# ---------------------------------------------------------------------------
# Audit Logon (bara lokalt och bara om Logon/FailedLogon ska sokas)
# ---------------------------------------------------------------------------
$needSuccess = $sources -contains 'Logon'
$needFailure = $sources -contains 'FailedLogon'
$auditOff = $false
$auditHelpShown = $false
$auditSetting = $null
if (($needSuccess -or $needFailure) -and -not $ComputerName -and (Test-IsAdministrator)) {
    $auditSetting = Get-AuditLogonSetting
    if ($null -eq $auditSetting) {
        Write-Host '  Audit Logon : kunde inte lasas (auditpol)'
    }
    else {
        $auditText = switch ($auditSetting) {
            0 { 'No Auditing (avstangd)' }
            1 { 'Success' }
            2 { 'Failure' }
            3 { 'Success och Failure' }
            default { "okant varde $auditSetting" }
        }
        Write-Host ("  Audit Logon : {0}" -f $auditText)

        $missing = @()
        if ($needSuccess -and -not ($auditSetting -band 1)) { $missing += 'Success -> 4624 loggas inte (kalla Logon)' }
        if ($needFailure -and -not ($auditSetting -band 2)) { $missing += 'Failure -> 4625 loggas inte (kalla FailedLogon)' }
        if ($missing.Count -gt 0) {
            $auditOff = $true
            Write-Host ''
            if ($auditSetting -eq 0) {
                Write-Warning 'Audit Logon ar avstangd pa den har datorn - nya inloggningar (4624/4625) skrivs inte till Security-loggen.'
            }
            else {
                Write-Warning 'Audit Logon ar avstangd for en del av det som efterfragas:'
            }
            foreach ($m in $missing) { Write-Host "    $m" -ForegroundColor Yellow }
            Show-AuditLogonHelp
            $auditHelpShown = $true
        }
    }
}
elseif (($needSuccess -or $needFailure) -and $ComputerName) {
    Write-Verbose 'Audit Logon kontrolleras bara pa den lokala datorn.'
}

# ---------------------------------------------------------------------------
# Loggdjup
# ---------------------------------------------------------------------------
if (-not $SkipLogInfo) {
    Write-Host "`n=== Loggdjup ===" -ForegroundColor Cyan
    foreach ($name in @($sources | ForEach-Object { $logNames[$_] } | Select-Object -Unique)) {
        Write-Host $name
        try {
            $p = @{ ListLog = $name; ErrorAction = 'Stop' }
            foreach ($k in $remote.Keys) { $p[$k] = $remote[$k] }
            $info = Get-WinEvent @p

            $p2 = @{ LogName = $name; Oldest = $true; MaxEvents = 1; ErrorAction = 'Stop' }
            foreach ($k in $remote.Keys) { $p2[$k] = $remote[$k] }
            $oldest = $null
            try { $oldest = Get-WinEvent @p2 } catch { $oldest = $null }

            if ($oldest) {
                $daysBack = [math]::Round(((Get-Date) - $oldest.TimeCreated).TotalDays, 1)
                Write-Host ("  Aldsta handelse : {0:yyyy-MM-dd HH:mm:ss}  ({1} dagar bakat)" -f $oldest.TimeCreated, $daysBack)
                if ($null -ne $from -and $oldest.TimeCreated -gt $from) {
                    Write-Warning ("Loggen borjar {0:yyyy-MM-dd HH:mm}, senare an begard start - perioden tackes inte helt." -f $oldest.TimeCreated)
                }
            }
            else {
                Write-Host '  Aldsta handelse : (loggen ar tom)'
            }
            Write-Host ("  Loggstorlek     : {0:N1} MB av max {1:N1} MB  (overwrite: {2})" -f ($info.FileSize / 1MB), ($info.MaximumSizeInBytes / 1MB), $info.LogMode)
            Write-Host ("  Antal handelser : {0}" -f $info.RecordCount)
            if (-not $info.IsEnabled) { Write-Warning "Loggen ar avstangd (IsEnabled = False)." }
        }
        catch {
            Write-Warning "Kunde inte lasa $name : $($_.Exception.Message)"
        }
    }
}

# ---------------------------------------------------------------------------
# Hamta handelser
# ---------------------------------------------------------------------------
Write-Host "`n=== Hamtar handelser ===" -ForegroundColor Cyan
$all = New-Object System.Collections.Generic.List[object]
$readCount = @{}

foreach ($s in $sources) {
    switch ($s) {
        'Logon'       { $xp = New-EventXPath -EventId 4624 -TypeIds $typeIds }
        'FailedLogon' {
            if ($typeFilterGiven) { $xp = New-EventXPath -EventId 4625 -TypeIds $typeIds }
            else { $xp = New-EventXPath -EventId 4625 }
        }
        'RdsSession'  {
            if ($IncludeReconnect) { $xp = New-EventXPath -EventId 21, 25 }
            else { $xp = New-EventXPath -EventId 21 }
        }
        'RdpAuth'     { $xp = New-EventXPath -EventId 1149 }
        'NpsAccess'   { $xp = New-EventXPath -EventId 6272 }
    }

    try {
        $events = @(Invoke-EventQuery -LogName $logNames[$s] -XPath $xp)
    }
    catch {
        Write-Warning "Kalla '$s' ($($logNames[$s])): $($_.Exception.Message)"
        continue
    }

    $kept = 0
    foreach ($e in $events) {
        $rec = ConvertTo-Record -Ev $e -Src $s
        if ($null -ne $rec -and (Test-UserWanted $rec.User)) {
            $all.Add($rec)
            $kept++
        }
    }
    $readCount[$s] = $events.Count
    Write-Host ("  {0,-12}: {1} handelser lasta, {2} efter filter" -f $s, $events.Count, $kept)
}

# 0 traffar i Logon/FailedLogon: forklara utifran vad auditkontrollen visade
$securityEmpty = @($sources | Where-Object { $_ -in @('Logon', 'FailedLogon') -and $readCount.ContainsKey($_) -and $readCount[$_] -eq 0 })
if ($securityEmpty.Count -gt 0) {
    $bitFor = @{ Logon = 1; FailedLogon = 2 }
    $emptyOff = @(); $emptyOn = @(); $emptyUnknown = @()
    foreach ($s in $securityEmpty) {
        if ($null -eq $auditSetting) { $emptyUnknown += $s }
        elseif ($auditSetting -band $bitFor[$s]) { $emptyOn += $s }
        else { $emptyOff += $s }
    }
    Write-Host ''
    if ($emptyOff.Count -gt 0) {
        Write-Warning ("{0}: 0 handelser - Audit Logon ar avstangd (se ovan)." -f ($emptyOff -join ', '))
    }
    if ($emptyOn.Count -gt 0) {
        Write-Warning ("{0}: 0 handelser. Audit Logon ar paslagen, sa det fanns inga traffar for vald period och valda inloggningstyper. Prova en langre period eller fler typer med -LogonType." -f ($emptyOn -join ', '))
    }
    if ($emptyUnknown.Count -gt 0) {
        Write-Warning ("{0}: 0 handelser i Security-loggen. Vanligaste orsaken: Audit Logon ar avstangd. Kontrollera pa datorn med:" -f ($emptyUnknown -join ', '))
        Write-Host ('    auditpol /get /subcategory:"{0}"' -f $auditLogonGuid)
        if (-not $auditHelpShown) { Show-AuditLogonHelp }
    }
}

if ($all.Count -eq 0) {
    Write-Warning 'Inga inloggningar hittades med valda kallor och filter.'
    return
}

# ---------------------------------------------------------------------------
# Sammanstall
# ---------------------------------------------------------------------------
$timeFmt = 'yyyy-MM-dd HH:mm:ss'

if ($Detailed) {
    $out = @(
        $all | Sort-Object Time | ForEach-Object {
            [pscustomobject]@{
                Tid      = $_.Time.ToString($timeFmt)
                Dator    = $_.Computer
                Kalla    = $_.Source
                Anvandare = $_.User
                IP       = $_.IP
                Detalj   = $_.Detail
            }
        }
    )
}
else {
    $rows = foreach ($g in ($all | Group-Object Source, User)) {
        $sorted = @($g.Group | Sort-Object Time)
        $firstRec = $sorted[0]
        $lastRec = $sorted[$sorted.Count - 1]
        [pscustomobject]@{
            Kalla         = $firstRec.Source
            Anvandare     = $firstRec.User
            Antal         = $sorted.Count
            Forsta        = $firstRec.Time.ToString($timeFmt)
            Senaste       = $lastRec.Time.ToString($timeFmt)
            SenasteIP     = $lastRec.IP
            SenasteDetalj = $lastRec.Detail
        }
    }
    $out = @($rows | Sort-Object Kalla, @{ Expression = 'Antal'; Descending = $true }, Anvandare)
}

Write-Host ''
if ($PassThru) {
    # Tabellen skrivs inte ut - objekten skickas vidare i pipelinen (sist i skriptet).
}
else {
    $out | Format-Table -AutoSize | Out-Host
}

Write-Host '--- Summering ---' -ForegroundColor Cyan
foreach ($s in $sources) {
    $r = @($all | Where-Object { $_.Source -eq $s })
    if ($r.Count -eq 0) {
        Write-Host ("{0,-12}: inga inloggningar" -f $s)
        continue
    }
    $users = @($r | Group-Object User).Count
    $m = $r | Measure-Object -Property Time -Minimum -Maximum
    Write-Host ("{0,-12}: {1} st, {2} unika anvandare, {3:yyyy-MM-dd} - {4:yyyy-MM-dd}" -f $s, $r.Count, $users, $m.Minimum, $m.Maximum)
}

if (-not $NoCsv) {
    if (-not $OutCsv) {
        $suffix = if ($Detailed) { '_detaljerad' } else { '' }
        $fileName = 'LogonReport_{0}_{1}{2}.csv' -f $target.Split('.')[0], (Get-Date -Format 'yyyyMMdd_HHmm'), $suffix
        $OutCsv = Join-Path (Get-Location).Path $fileName
    }
    $out | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host "Sparad: $OutCsv" -ForegroundColor Green
}

if ($PassThru) { $out }
