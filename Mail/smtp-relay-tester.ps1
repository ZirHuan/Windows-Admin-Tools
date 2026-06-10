#Requires -Version 5.1

<#
.SYNOPSIS
    Interactive SMTP relay diagnostic and test tool.

.DESCRIPTION
    Connects to an SMTP relay over a raw TCP socket (TcpClient + SslStream),
    runs EHLO, inspects advertised capabilities and AUTH methods, optionally
    authenticates (AUTH LOGIN or AUTH PLAIN), and optionally sends a test
    message.  The full SMTP protocol conversation is shown in verbose mode.

    Equivalent to smtp-relay-tester.py but implemented entirely in PowerShell
    over a raw socket so that every server response line is visible and the
    EHLO capability list can be inspected directly.

    Last-used settings (including DPAPI/AES-encrypted password) are persisted
    to $HOME/.smtp-relay-tester.json so subsequent runs pre-fill all prompts.

    Compatible with Windows PowerShell 5.1 and PowerShell 7+.

.PARAMETER SettingsFile
    Path to the JSON settings file.  Defaults to $HOME/.smtp-relay-tester.json.

.EXAMPLE
    powershell -File smtp-relay-tester.ps1

.EXAMPLE
    pwsh -File smtp-relay-tester.ps1

.EXAMPLE
    pwsh -File smtp-relay-tester.ps1 -SettingsFile C:\Temp\myrelay.json

.NOTES
    Compatible with Windows PowerShell 5.1 and PowerShell 7+.
    On Windows, password is encrypted with DPAPI (per-user, per-machine).
    On Linux/macOS (PS7 only), a random AES-256 key is stored alongside the
    settings file as a .key file.  Neither mechanism protects against an
    attacker who has access to the running user account.
#>

[CmdletBinding()]
param(
    [string]$SettingsFile = (Join-Path ([Environment]::GetFolderPath('UserProfile')) '.smtp-relay-tester.json')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Tool version - shown in the banner and tracked in CHANGELOG.md
$ScriptVersion = '1.0.0'

# ---------------------------------------------------------------------------
# PS 5.1 / 7 compatibility shims for automatic platform variables
# $IsWindows, $IsLinux, $IsMacOS do not exist in Windows PowerShell 5.1.
# Under Set-StrictMode -Version Latest, referencing a non-existent variable
# throws, so we define safe script-scope aliases here.
# ---------------------------------------------------------------------------
if (Get-Variable -Name IsWindows -ErrorAction SilentlyContinue) {
    $script:IsWindowsCompat = $IsWindows
    $script:IsLinuxCompat   = $IsLinux
    $script:IsMacOSCompat   = $IsMacOS
}
else {
    # Windows PowerShell 5.1 only runs on Windows
    $script:IsWindowsCompat = $true
    $script:IsLinuxCompat   = $false
    $script:IsMacOSCompat   = $false
}

# ---------------------------------------------------------------------------
# Helper: safe dictionary lookup with fallback (replaces PS7-only ?? operator)
# Works with both [hashtable] and [System.Collections.Specialized.OrderedDictionary].
# ---------------------------------------------------------------------------
function Get-OrDefault {
    param(
        [System.Collections.IDictionary]$Hash,
        [string]$Key,
        [object]$Default = ''
    )
    if ($null -ne $Hash -and $Hash.Contains($Key) -and $null -ne $Hash[$Key] -and $Hash[$Key] -ne '') {
        return $Hash[$Key]
    }
    return $Default
}

# ---------------------------------------------------------------------------
# Output helpers
# ---------------------------------------------------------------------------
function Write-Status {
    param(
        [ValidateSet('OK','WARN','FAIL','INFO','PROTO')]
        [string]$Level,
        [string]$Message
    )
    $prefix = switch ($Level) {
        'OK'    { '[OK]  ' }
        'WARN'  { '[WARN] ' }
        'FAIL'  { '[FAIL] ' }
        'INFO'  { '[INFO] ' }
        'PROTO' { '[....] ' }
    }
    $color = switch ($Level) {
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'FAIL'  { 'Red' }
        'INFO'  { 'Cyan' }
        'PROTO' { 'DarkGray' }
    }
    Write-Host "$prefix$Message" -ForegroundColor $color
}

# Only writes when $script:VerboseMode is true
function Write-Proto {
    param([string]$Direction, [string]$Line)
    if ($script:VerboseMode) {
        $arrow = if ($Direction -eq 'C') { 'C>>>' } else { 'S<<<' }
        Write-Host "  $arrow $Line" -ForegroundColor DarkGray
    }
}

# ---------------------------------------------------------------------------
# Settings persistence
# ---------------------------------------------------------------------------
$KeyFile = $SettingsFile + '.key'

function Get-AesKey {
    <#
    Returns a 32-byte AES key for Linux/macOS (DPAPI is not available there).
    The key is stored alongside the settings file.  If the key file does not
    exist yet, a new random key is generated and saved.
    This function is only called on non-Windows (PS7+).
    #>
    if (Test-Path $KeyFile) {
        $bytes = [System.IO.File]::ReadAllBytes($KeyFile)
        if ($bytes.Length -eq 32) { return $bytes }
    }
    $key = [byte[]]::new(32)
    # Use GetBytes() for cross-version compatibility (Fill() is .NET Core only)
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $rng.GetBytes($key)
    $rng.Dispose()
    [System.IO.File]::WriteAllBytes($KeyFile, $key)
    # Restrict permissions on Linux/macOS
    if ($script:IsLinuxCompat -or $script:IsMacOSCompat) {
        & chmod 600 $KeyFile 2>$null
    }
    return $key
}

function Protect-Password {
    param([securestring]$SecurePassword)
    # Returns a base64 string safe for JSON storage
    if ($script:IsWindowsCompat) {
        # DPAPI: ConvertFrom-SecureString with no Key uses per-user Windows DPAPI
        return ConvertFrom-SecureString -SecureString $SecurePassword
    }
    else {
        # AES-256 with a per-file random key stored next to the settings file
        $key = Get-AesKey
        return ConvertFrom-SecureString -SecureString $SecurePassword -Key $key
    }
}

function Unprotect-Password {
    param([string]$EncryptedString)
    # Returns a SecureString, or $null on failure
    try {
        if ($script:IsWindowsCompat) {
            return ConvertTo-SecureString -String $EncryptedString -ErrorAction Stop
        }
        else {
            if (-not (Test-Path $KeyFile)) { return $null }
            $key = Get-AesKey
            return ConvertTo-SecureString -String $EncryptedString -Key $key -ErrorAction Stop
        }
    }
    catch {
        return $null
    }
}

function Load-Settings {
    if (-not (Test-Path $SettingsFile)) { return @{} }
    try {
        $raw = Get-Content $SettingsFile -Raw
        # -AsHashtable was added in PS6; use a manual conversion on 5.1
        if ($PSVersionTable.PSVersion.Major -ge 6) {
            $result = $raw | ConvertFrom-Json -AsHashtable
        }
        else {
            $pso = $raw | ConvertFrom-Json
            $result = @{}
            foreach ($prop in $pso.PSObject.Properties) {
                $result[$prop.Name] = $prop.Value
            }
        }
        return $result
    }
    catch {
        Write-Status INFO "Could not read settings file ($($_.Exception.Message)) - starting fresh."
        return @{}
    }
}

function Save-Settings {
    # Accept both [hashtable] and [ordered]@{} (OrderedDictionary) -
    # IDictionary is the common base type for both.
    param([System.Collections.IDictionary]$Settings)
    try {
        $Settings | ConvertTo-Json -Depth 3 | Set-Content -Path $SettingsFile -Encoding UTF8
        if ($script:IsLinuxCompat -or $script:IsMacOSCompat) {
            & chmod 600 $SettingsFile 2>$null
        }
    }
    catch {
        Write-Status WARN "Could not save settings: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Prompt helpers that accept a bracketed default on Enter
# ---------------------------------------------------------------------------
function Prompt-WithDefault {
    param(
        [string]$Prompt,
        [string]$Default = ''
    )
    # Wrap $Prompt in braces so PS does not read "$Prompt:" as a
    # scope-qualified variable reference (parser crash on PS7).
    # Note: Read-Host appends its own ": ", so we do NOT add one here.
    $display = if ($Default -ne '') { "${Prompt} [${Default}]" } else { "${Prompt}" }
    $answer = Read-Host -Prompt $display
    if ($answer -eq '') { return $Default }
    return $answer
}

function Prompt-YesNo {
    param(
        [string]$Prompt,
        [bool]$Default = $true
    )
    $hint = if ($Default) { 'Y/n' } else { 'y/N' }
    $reply = Read-Host -Prompt "$Prompt [$hint]"
    if ($reply -eq '') { return $Default }
    return $reply.Trim().ToLower() -in @('y','yes')
}

function Prompt-Password {
    param(
        [string]$Prompt = 'Password',
        [bool]$HasSavedPassword = $false
    )
    # If there is a saved password, offer it as a masked default
    if ($HasSavedPassword) {
        $reply = Read-Host -Prompt "$Prompt [******** (saved, Enter to keep)]"
        if ($reply -eq '') { return $null }   # caller keeps existing SecureString
        return (ConvertTo-SecureString -String $reply -AsPlainText -Force)
    }
    else {
        $ss = Read-Host -Prompt $Prompt -AsSecureString
        return $ss
    }
}

# ---------------------------------------------------------------------------
# Port picker
# ---------------------------------------------------------------------------
# NOTE: keys are STRINGS on purpose. An [ordered] dictionary (OrderedDictionary)
# indexed by an [int] uses POSITIONAL indexing, not key lookup - so integer keys
# collide with positions and throw "Index out of range". String keys force
# key-based lookup.
$PortMenu = [ordered]@{
    '1' = @{ Port = 25;   Label = 'Plain SMTP (no encryption)' }
    '2' = @{ Port = 587;  Label = 'SMTP + STARTTLS (submission, recommended)' }
    '3' = @{ Port = 465;  Label = 'SMTP over implicit SSL/TLS' }
    '4' = @{ Port = 2525; Label = 'Common alternate submission port' }
    '5' = @{ Port = 0;    Label = 'Custom port' }
}

function Prompt-Port {
    param([int]$DefaultPort = 587)
    Write-Host ""
    Write-Host "  Select SMTP port:" -ForegroundColor White
    foreach ($k in $PortMenu.Keys) {
        $entry = $PortMenu[$k]
        $marker = if ($entry.Port -eq $DefaultPort) { '*' } else { ' ' }
        if ($entry.Port -eq 0) {
            Write-Host ("  [{0}]{1} {2,-6}  {3}" -f $k, $marker, 'custom', $entry.Label)
        }
        else {
            Write-Host ("  [{0}]{1} {2,-6}  {3}" -f $k, $marker, $entry.Port, $entry.Label)
        }
    }
    Write-Host ""

    # Find which menu number corresponds to the default port
    $defaultChoice = ''
    foreach ($k in $PortMenu.Keys) {
        if ($PortMenu[$k].Port -eq $DefaultPort) { $defaultChoice = "$k"; break }
    }
    # If the saved default is a custom port, pre-fill the raw number
    if ($defaultChoice -eq '') { $defaultChoice = "$DefaultPort" }

    $choice = Read-Host -Prompt "  Choice or port number [$defaultChoice]"
    if ($choice -eq '') { $choice = $defaultChoice }
    $choice = $choice.Trim()

    # Menu selection: keys are strings '1'..'5', so match the raw string
    # (string key lookup avoids the OrderedDictionary positional-index trap).
    if ($PortMenu.Contains($choice)) {
        if ($choice -eq '5') {
            # Custom port
            $custom = Read-Host -Prompt "  Enter custom port"
            $customInt = 0
            if ([int]::TryParse($custom, [ref]$customInt) -and $customInt -gt 0 -and $customInt -le 65535) {
                return $customInt
            }
            Write-Status WARN "Invalid port; using $DefaultPort"
            return $DefaultPort
        }
        return $PortMenu[$choice].Port
    }

    # Otherwise treat the input as a raw port number typed directly
    $asInt = 0
    if ([int]::TryParse($choice, [ref]$asInt) -and $asInt -gt 0 -and $asInt -le 65535) {
        return $asInt
    }

    Write-Status WARN "Could not parse port; using $DefaultPort"
    return $DefaultPort
}

# ---------------------------------------------------------------------------
# Raw SMTP conversation helpers
# ---------------------------------------------------------------------------
function Read-SmtpLine {
    param([System.IO.StreamReader]$Reader)
    $line = $Reader.ReadLine()
    Write-Proto 'S' $line
    return $line
}

function Read-SmtpResponse {
    <#
    Reads a (possibly multi-line) SMTP response.
    Returns a hashtable: @{ Code = '250'; Lines = @('250-EHLO...','250 OK') }
    #>
    param([System.IO.StreamReader]$Reader)
    $lines = [System.Collections.Generic.List[string]]::new()
    do {
        $line = Read-SmtpLine $Reader
        if ($null -eq $line) { break }
        $lines.Add($line)
        # Multi-line responses have a '-' at position 3; single/last line has ' '
    } while ($line.Length -ge 4 -and $line[3] -eq '-')

    $code = if ($lines.Count -gt 0 -and $lines[0].Length -ge 3) { $lines[0].Substring(0,3) } else { '000' }
    return @{ Code = $code; Lines = $lines }
}

function Send-SmtpCommand {
    param(
        [System.IO.StreamWriter]$Writer,
        [string]$Command,
        [bool]$Redact = $false
    )
    $display = if ($Redact) { ($Command -replace '(?<=\s)\S+$','********') } else { $Command }
    Write-Proto 'C' $display
    $Writer.WriteLine($Command)
    $Writer.Flush()
}

function Assert-SmtpCode {
    param(
        [hashtable]$Response,
        [string[]]$Expected,
        [string]$Context
    )
    if ($Response.Code -notin $Expected) {
        throw "SMTP error during ${Context}: expected $($Expected -join '/'), got $($Response.Code) - $($Response.Lines[-1])"
    }
}

# ---------------------------------------------------------------------------
# TLS upgrade helper
# ---------------------------------------------------------------------------
function Upgrade-ToTls {
    param(
        [System.Net.Sockets.TcpClient]$Tcp,
        [string]$ServerName,
        [ref]$ReaderRef,
        [ref]$WriterRef
    )
    $netStream = $Tcp.GetStream()
    $sslStream = [System.Net.Security.SslStream]::new(
        $netStream,
        $false,   # leave inner stream open
        { param($s,$cert,$chain,$err) $true }   # accept all certs for diagnostic tool
    )
    $sslStream.AuthenticateAsClient($ServerName)

    if ($script:VerboseMode) {
        Write-Host ("  [TLS] Protocol : {0}" -f $sslStream.SslProtocol) -ForegroundColor DarkGreen
        Write-Host ("  [TLS] Cipher   : {0}" -f $sslStream.CipherAlgorithm) -ForegroundColor DarkGreen
        Write-Host ("  [TLS] Strength : {0} bits" -f $sslStream.CipherStrength) -ForegroundColor DarkGreen
        Write-Host ("  [TLS] Hash     : {0}" -f $sslStream.HashAlgorithm) -ForegroundColor DarkGreen
        $cert = $sslStream.RemoteCertificate
        if ($cert) {
            Write-Host ("  [TLS] Cert CN  : {0}" -f $cert.Subject) -ForegroundColor DarkGreen
            Write-Host ("  [TLS] Cert exp : {0}" -f $cert.GetExpirationDateString()) -ForegroundColor DarkGreen
        }
    }

    $encoding = [System.Text.Encoding]::ASCII
    $ReaderRef.Value = [System.IO.StreamReader]::new($sslStream, $encoding)
    $WriterRef.Value = [System.IO.StreamWriter]::new($sslStream, $encoding)
    $WriterRef.Value.NewLine = "`r`n"
    $WriterRef.Value.AutoFlush = $false
}

# ---------------------------------------------------------------------------
# Parse EHLO capabilities from the multi-line response
# ---------------------------------------------------------------------------
function Parse-EhloCapabilities {
    param([hashtable]$EhloResponse)
    $caps = [ordered]@{}
    foreach ($line in $EhloResponse.Lines) {
        if ($line.Length -lt 4) { continue }
        $body = $line.Substring(4).Trim()
        if ($body -eq '') { continue }
        $parts = $body.Split(' ', 2)
        $key = $parts[0].ToUpper()
        $val = if ($parts.Count -gt 1) { $parts[1] } else { '' }
        $caps[$key] = $val
    }
    return $caps
}

# ---------------------------------------------------------------------------
# AUTH LOGIN implementation
# ---------------------------------------------------------------------------
function Invoke-AuthLogin {
    param(
        [System.IO.StreamReader]$Reader,
        [System.IO.StreamWriter]$Writer,
        [string]$Username,
        [securestring]$SecurePassword
    )
    Send-SmtpCommand $Writer 'AUTH LOGIN'
    $r1 = Read-SmtpResponse $Reader
    Assert-SmtpCode $r1 @('334') 'AUTH LOGIN challenge (username)'

    $b64User = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($Username))
    Send-SmtpCommand $Writer $b64User
    $r2 = Read-SmtpResponse $Reader
    Assert-SmtpCode $r2 @('334') 'AUTH LOGIN challenge (password)'

    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    try {
        $plain = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        $b64Pass = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($plain))
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
    Send-SmtpCommand $Writer $b64Pass -Redact $true
    $r3 = Read-SmtpResponse $Reader
    Assert-SmtpCode $r3 @('235') 'AUTH LOGIN (password accepted)'
}

# ---------------------------------------------------------------------------
# AUTH PLAIN implementation
# ---------------------------------------------------------------------------
function Invoke-AuthPlain {
    param(
        [System.IO.StreamReader]$Reader,
        [System.IO.StreamWriter]$Writer,
        [string]$Username,
        [securestring]$SecurePassword
    )
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    try {
        $plain = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        # AUTH PLAIN format: NUL + username + NUL + password
        $authBytes = [System.Text.Encoding]::UTF8.GetBytes("`0$Username`0$plain")
        $b64Auth = [Convert]::ToBase64String($authBytes)
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
    Send-SmtpCommand $Writer "AUTH PLAIN $b64Auth" -Redact $true
    $r = Read-SmtpResponse $Reader
    Assert-SmtpCode $r @('235') 'AUTH PLAIN'
}

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

$script:VerboseMode = $false

Write-Host ""
Write-Host "=== SMTP Relay Tester v$ScriptVersion ===" -ForegroundColor White
Write-Host "  PS 5.1/7 edition - raw socket SMTP diagnostic" -ForegroundColor DarkCyan
Write-Host ""

# Load persisted settings
$saved = Load-Settings

# ---- Gather connection parameters ----

$server = Prompt-WithDefault 'SMTP server hostname' (Get-OrDefault $saved 'server' '')
if ($server -eq '') {
    Write-Status FAIL "Server hostname is required."
    exit 1
}

$savedPort = if ($saved.Contains('port') -and $null -ne $saved['port']) { [int]$saved['port'] } else { 587 }
$port = Prompt-Port -DefaultPort $savedPort

$savedVerbose = if ($saved.Contains('verbose') -and $null -ne $saved['verbose']) { [bool]$saved['verbose'] } else { $false }
$verboseMode = Prompt-YesNo 'Verbose output (show full SMTP conversation)' $savedVerbose
$script:VerboseMode = $verboseMode

$savedDoAuth = if ($saved.Contains('doAuth') -and $null -ne $saved['doAuth']) { [bool]$saved['doAuth'] } else { $true }
$doAuth = Prompt-YesNo 'Authenticate' $savedDoAuth

$username = ''
$securePassword = $null
$hadSavedPassword = $false

if ($doAuth) {
    $username = Prompt-WithDefault 'Username' (Get-OrDefault $saved 'username' '')

    # Attempt to restore saved password
    $savedEncPwd = Get-OrDefault $saved 'encryptedPassword' ''
    $restoredPwd = $null
    if ($savedEncPwd -ne '') {
        $restoredPwd = Unprotect-Password $savedEncPwd
        if ($null -ne $restoredPwd) {
            $hadSavedPassword = $true
        }
        else {
            Write-Status WARN "Saved password could not be decrypted (different machine/user?) - please re-enter."
        }
    }

    $newPwd = Prompt-Password -Prompt 'Password' -HasSavedPassword $hadSavedPassword
    if ($null -eq $newPwd -and $hadSavedPassword) {
        # User pressed Enter - keep the restored password
        $securePassword = $restoredPwd
    }
    elseif ($null -ne $newPwd) {
        $securePassword = $newPwd
    }
    else {
        Write-Status FAIL "Password is required when authentication is enabled."
        exit 1
    }
}

$defaultFrom = if ($username -ne '') { "$username@$server" } else { "test@$server" }
$savedFrom = Get-OrDefault $saved 'fromAddress' $defaultFrom
$fromAddress = Prompt-WithDefault 'From address' $savedFrom

$savedTo = Get-OrDefault $saved 'toAddress' ''
$toAddress = Prompt-WithDefault 'Send test mail to (Enter to skip)' $savedTo

Write-Host ""

# ---- Determine TLS mode from port ----
# 465  -> implicit SSL (wrap socket in SslStream before banner)
# 587  -> STARTTLS after EHLO
# other-> plain unless server advertises STARTTLS (then upgrade opportunistically)
$tlsMode = switch ($port) {
    465     { 'ImplicitSSL' }
    587     { 'STARTTLS' }
    default { 'OpportunisticSTARTTLS' }
}

Write-Status INFO "Connecting to ${server}:${port} (TLS mode: $tlsMode) ..."

$tcp = $null
$reader = $null
$writer = $null
$exitCode = 0

try {
    # ---- TCP connect ----
    $tcp = [System.Net.Sockets.TcpClient]::new()
    $connectTask = $tcp.ConnectAsync($server, $port)
    if (-not $connectTask.Wait(15000)) {
        throw "TCP connection timed out after 15 seconds."
    }
    if (-not $tcp.Connected) {
        throw "TCP connection failed (no exception detail available)."
    }
    Write-Status OK "TCP connection established to ${server}:${port}"

    $encoding = [System.Text.Encoding]::ASCII

    if ($tlsMode -eq 'ImplicitSSL') {
        # Wrap in SSL immediately before reading the banner
        $sslStreamObj = [System.Net.Security.SslStream]::new(
            $tcp.GetStream(),
            $false,
            { param($s,$cert,$chain,$err) $true }
        )
        $sslStreamObj.AuthenticateAsClient($server)
        if ($verboseMode) {
            Write-Host ("  [TLS] Protocol : {0}" -f $sslStreamObj.SslProtocol) -ForegroundColor DarkGreen
            Write-Host ("  [TLS] Cipher   : {0}" -f $sslStreamObj.CipherAlgorithm) -ForegroundColor DarkGreen
            Write-Host ("  [TLS] Strength : {0} bits" -f $sslStreamObj.CipherStrength) -ForegroundColor DarkGreen
            $cert = $sslStreamObj.RemoteCertificate
            if ($cert) {
                Write-Host ("  [TLS] Cert CN  : {0}" -f $cert.Subject) -ForegroundColor DarkGreen
                Write-Host ("  [TLS] Cert exp : {0}" -f $cert.GetExpirationDateString()) -ForegroundColor DarkGreen
            }
        }
        $reader = [System.IO.StreamReader]::new($sslStreamObj, $encoding)
        $writer = [System.IO.StreamWriter]::new($sslStreamObj, $encoding)
    }
    else {
        $netStream = $tcp.GetStream()
        $reader = [System.IO.StreamReader]::new($netStream, $encoding)
        $writer = [System.IO.StreamWriter]::new($netStream, $encoding)
    }

    $writer.NewLine    = "`r`n"
    $writer.AutoFlush  = $false

    # ---- Read server banner ----
    $banner = Read-SmtpResponse $reader
    Assert-SmtpCode $banner @('220') 'server banner'
    Write-Status OK "Server banner: $($banner.Lines[-1])"

    # ---- First EHLO ----
    $fqdn = [System.Net.Dns]::GetHostName()
    Send-SmtpCommand $writer "EHLO $fqdn"
    $ehloResp = Read-SmtpResponse $reader
    Assert-SmtpCode $ehloResp @('250') 'EHLO'

    $caps = Parse-EhloCapabilities $ehloResp

    # ---- STARTTLS if required ----
    if ($tlsMode -eq 'STARTTLS') {
        if (-not $caps.Contains('STARTTLS')) {
            throw "Port 587 selected but server did not advertise STARTTLS. Check server config."
        }
        Send-SmtpCommand $writer 'STARTTLS'
        $stResp = Read-SmtpResponse $reader
        Assert-SmtpCode $stResp @('220') 'STARTTLS'
        Write-Status OK "STARTTLS negotiating TLS ..."

        Upgrade-ToTls -Tcp $tcp -ServerName $server -ReaderRef ([ref]$reader) -WriterRef ([ref]$writer)
        Write-Status OK "TLS channel established"

        # Re-EHLO over TLS
        Send-SmtpCommand $writer "EHLO $fqdn"
        $ehloResp2 = Read-SmtpResponse $reader
        Assert-SmtpCode $ehloResp2 @('250') 'EHLO (post-TLS)'
        $caps = Parse-EhloCapabilities $ehloResp2
    }
    elseif ($tlsMode -eq 'OpportunisticSTARTTLS' -and $caps.Contains('STARTTLS')) {
        Write-Status INFO "Server advertises STARTTLS on port $port - upgrading opportunistically ..."
        Send-SmtpCommand $writer 'STARTTLS'
        $stResp2 = Read-SmtpResponse $reader
        Assert-SmtpCode $stResp2 @('220') 'opportunistic STARTTLS'
        Upgrade-ToTls -Tcp $tcp -ServerName $server -ReaderRef ([ref]$reader) -WriterRef ([ref]$writer)
        Write-Status OK "Opportunistic TLS established"
        Send-SmtpCommand $writer "EHLO $fqdn"
        $ehloResp3 = Read-SmtpResponse $reader
        Assert-SmtpCode $ehloResp3 @('250') 'EHLO (post-opportunistic-TLS)'
        $caps = Parse-EhloCapabilities $ehloResp3
    }

    # ---- Print capability table ----
    Write-Host ""
    Write-Host "  --- EHLO capabilities ---" -ForegroundColor White
    foreach ($k in $caps.Keys) {
        $v = $caps[$k]
        Write-Host ("  {0,-20} {1}" -f $k, $(if ($v -ne '') { $v } else { '(present)' })) -ForegroundColor Gray
    }
    Write-Host "  -------------------------" -ForegroundColor White
    Write-Host ""

    # ---- AUTH methods analysis ----
    if ($caps.Contains('AUTH')) {
        $authMethods = $caps['AUTH'].ToUpper().Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
        Write-Status INFO "Advertised AUTH methods: $($authMethods -join ', ')"
        if ('NTLM' -in $authMethods) {
            Write-Status WARN "NTLM is advertised - this breaks Veeam 13 / .NET 8 SmtpClient."
            Write-Host "         Ask the relay provider to remove NTLM, leaving only LOGIN and/or PLAIN." -ForegroundColor Yellow
        }
        else {
            Write-Status OK "NTLM not advertised - AUTH methods look fine."
        }
    }
    else {
        Write-Status INFO "No AUTH methods advertised in EHLO."
    }
    Write-Host ""

    # ---- Authentication ----
    if ($doAuth) {
        Write-Status INFO "Authenticating as '$username' ..."

        $availableAuth = @()
        if ($caps.Contains('AUTH')) {
            $availableAuth = $caps['AUTH'].ToUpper().Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
        }

        if ('LOGIN' -in $availableAuth) {
            Invoke-AuthLogin -Reader $reader -Writer $writer -Username $username -SecurePassword $securePassword
        }
        elseif ('PLAIN' -in $availableAuth) {
            Invoke-AuthPlain -Reader $reader -Writer $writer -Username $username -SecurePassword $securePassword
        }
        elseif ($availableAuth.Count -eq 0) {
            # Server may require auth but did not advertise it (or we're on plain SMTP) - try LOGIN anyway
            Write-Status WARN "No AUTH methods in EHLO; attempting AUTH LOGIN anyway ..."
            Invoke-AuthLogin -Reader $reader -Writer $writer -Username $username -SecurePassword $securePassword
        }
        else {
            throw "No supported AUTH method available. Server offers: $($availableAuth -join ', '). This tool supports LOGIN and PLAIN."
        }

        Write-Status OK "Authenticated successfully."
    }
    else {
        Write-Status INFO "Skipped authentication (not requested)."
    }

    # ---- Send test message ----
    if ($toAddress -ne '') {
        Write-Status INFO "Sending test message from <$fromAddress> to <$toAddress> ..."

        Send-SmtpCommand $writer "MAIL FROM:<$fromAddress>"
        $mfResp = Read-SmtpResponse $reader
        Assert-SmtpCode $mfResp @('250') 'MAIL FROM'

        Send-SmtpCommand $writer "RCPT TO:<$toAddress>"
        $rtResp = Read-SmtpResponse $reader
        Assert-SmtpCode $rtResp @('250','251') 'RCPT TO'

        Send-SmtpCommand $writer 'DATA'
        $dataResp = Read-SmtpResponse $reader
        Assert-SmtpCode $dataResp @('354') 'DATA'

        $timestamp = (Get-Date -Format 'ddd, dd MMM yyyy HH:mm:ss zzz')
        $msgLines = @(
            "From: $fromAddress",
            "To: $toAddress",
            "Subject: SMTP Relay Test",
            "Date: $timestamp",
            "MIME-Version: 1.0",
            "Content-Type: text/plain; charset=UTF-8",
            "",
            "This is a test message from smtp-relay-tester.ps1.",
            "Server : ${server}:${port}",
            "User   : $username",
            "Time   : $timestamp",
            "."
        )
        foreach ($msgLine in $msgLines) {
            Write-Proto 'C' $msgLine
            $writer.WriteLine($msgLine)
        }
        $writer.Flush()

        $dotResp = Read-SmtpResponse $reader
        Assert-SmtpCode $dotResp @('250') 'message accepted'
        Write-Status OK "Test message accepted by server for delivery to <$toAddress>"
    }
    else {
        Write-Status INFO "No To address supplied - skipping test send."
    }

    # ---- QUIT ----
    Send-SmtpCommand $writer 'QUIT'
    $quitResp = Read-SmtpResponse $reader
    # 221 is normal; tolerate anything for a graceful shutdown
    Write-Status OK "Session closed (server: $($quitResp.Lines[-1]))"

    Write-Host ""
    Write-Status OK "Done - relay is working."
}
catch {
    Write-Host ""
    Write-Status FAIL "$($_.Exception.Message)"
    if ($verboseMode -and $_.ScriptStackTrace) {
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    }
    $exitCode = 1
}
finally {
    # Clean up streams and TCP connection
    try { if ($writer) { $writer.Close() } } catch {}
    try { if ($reader) { $reader.Close() } } catch {}
    try { if ($tcp)    { $tcp.Close()   } } catch {}

    # ---- Persist settings ----
    $settingsToSave = [ordered]@{
        server            = $server
        port              = $port
        username          = $username
        doAuth            = $doAuth
        fromAddress       = $fromAddress
        toAddress         = $toAddress
        verbose           = $verboseMode
        encryptedPassword = ''
    }

    if ($doAuth -and $null -ne $securePassword) {
        try {
            $settingsToSave['encryptedPassword'] = Protect-Password $securePassword
        }
        catch {
            Write-Status WARN "Could not encrypt password for storage: $($_.Exception.Message)"
        }
    }

    Save-Settings $settingsToSave
    Write-Status INFO "Settings saved to: $SettingsFile"
}

exit $exitCode
