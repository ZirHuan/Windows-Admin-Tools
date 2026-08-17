#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Creates an encrypted SMTP credential store for ServiceMonitor.ps1.

.DESCRIPTION
    Generates a random AES-256 key, prompts for the SMTP password,
    encrypts it using ConvertFrom-SecureString with the AES key, and saves:

        smtp.key   - the AES-256 key (32 bytes, binary)
        smtp.cred  - the AES-encrypted password (base64 text)
        smtp.user  - the SMTP username (plaintext - usernames are not secret)
        smtp.json  - per-deployment SMTP settings: server, port, from, user, useSsl
                     (no secrets). ServiceMonitor.ps1 auto-loads this from next to
                     the cred file, so the relay can differ per server.

    The .key file is protected with NTFS ACLs (SYSTEM + Administrators only).
    Anyone who obtains the .key file can decrypt the password - guard it accordingly.

    Pass the resulting files to ServiceMonitor.ps1:
        -SmtpUser (Get-Content smtp.user) -SmtpKeyFile .\smtp.key -SmtpCredFile .\smtp.cred

    Run this script once, then keep the files in a protected folder.
    Re-run only when the SMTP password changes.

.PARAMETER OutputFolder
    Where to save smtp.key, smtp.cred, and smtp.user.
    Default: same folder as this script.

.PARAMETER SmtpUser
    The SMTP username. If omitted, you will be prompted.
    Saved to smtp.user (plaintext).

.PARAMETER Force
    Overwrite existing credential files without prompting.

.EXAMPLE
    .\New-CredStore.ps1

.EXAMPLE
    .\New-CredStore.ps1 -SmtpUser relay@example.com -OutputFolder C:\Monitoring\creds

.NOTES
    The AES key is generated with RandomNumberGenerator (cryptographically secure).
    Encryption uses ConvertFrom-SecureString with an explicit 256-bit key (cross-platform;
    does NOT use DPAPI, so the credential files are portable between machines that
    share the same .key file).
    Requires elevation: the smtp.key ACL restriction (SYSTEM + Administrators)
    and the usual C:\ServiceMonitor target folder both need admin rights.
    Version: 1.3.0
#>

[CmdletBinding()]
param(
    [string] $OutputFolder = '',
    [string] $SmtpUser     = '',
    [string] $SmtpServer   = '',
    [int]    $SmtpPort     = 587,
    [string] $FromAddress  = '',
    [switch] $SmtpUseSsl,
    [switch] $Force
)

$ErrorActionPreference = 'Stop'

# $PSScriptRoot is empty inside param() when launched via 'powershell.exe -File'
# - resolve at body scope instead of defaulting the parameter.
if (-not $OutputFolder) {
    $OutputFolder = if     ($PSScriptRoot)  { $PSScriptRoot }
                    elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath }
                    else { throw 'Cannot resolve the script folder - pass -OutputFolder explicitly.' }
}

$keyFile  = Join-Path $OutputFolder 'smtp.key'
$credFile = Join-Path $OutputFolder 'smtp.cred'
$userFile = Join-Path $OutputFolder 'smtp.user'
$jsonFile = Join-Path $OutputFolder 'smtp.json'

# Guard against accidental overwrite
if (-not $Force) {
    foreach ($f in @($keyFile, $credFile, $userFile, $jsonFile)) {
        if (Test-Path -LiteralPath $f) {
            $answer = Read-Host "File exists: $f  Overwrite? [y/N]"
            if ($answer -notmatch '^[Yy]') {
                Write-Host 'Aborted.' -ForegroundColor Yellow
                exit 0
            }
        }
    }
}

if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

# Collect inputs
if (-not $SmtpUser) {
    $SmtpUser = Read-Host 'SMTP username'
}
if (-not $SmtpUser) {
    throw 'SMTP username cannot be empty.'
}

# SMTP server differs per deployment, so capture it here and store it alongside
# the credential (in smtp.json) rather than relying on a script default.
if (-not $SmtpServer) {
    $SmtpServer = Read-Host 'SMTP server (hostname or IP, WITHOUT port)'
}
if (-not $SmtpServer) {
    throw 'SMTP server cannot be empty.'
}
# host:port is a classic paste mistake - the whole string would be treated as a
# hostname by SmtpClient and DNS resolution would fail. Port goes in SmtpPort.
if ($SmtpServer -match ':') {
    throw "SMTP server must be a hostname or IP only (got '$SmtpServer'). Set the port separately - it is prompted next / use -SmtpPort."
}

# Port, SSL and From were previously silent defaults (25 / off / empty), which
# cost a real deployment hours: mail either never connected or was accepted by
# the relay and then dropped because the From domain was not authorized.
# Prompt for all three unless given on the command line.
if (-not $PSBoundParameters.ContainsKey('SmtpPort')) {
    $portAnswer = Read-Host "SMTP port [$SmtpPort]"
    if ($portAnswer) {
        # Validate before casting: '[int]' on junk throws an unfriendly parse
        # error, and 0/70000 would be written to smtp.json unchallenged.
        if ($portAnswer -notmatch '^\d+$' -or [int]$portAnswer -lt 1 -or [int]$portAnswer -gt 65535) {
            throw "Invalid SMTP port '$portAnswer' - must be a number 1-65535 (587 for STARTTLS relay, 25 for anonymous)."
        }
        $SmtpPort = [int]$portAnswer
    }
}
if (-not $PSBoundParameters.ContainsKey('SmtpUseSsl')) {
    $sslAnswer = Read-Host 'Use SSL/STARTTLS? (required on port 587) [Y/n]'
    $SmtpUseSsl = [switch]($sslAnswer -notmatch '^[Nn]')
}
if (-not $FromAddress) {
    $FromAddress = Read-Host 'From address (MUST be on a domain the relay accepts mail for)'
}
if (-not $FromAddress) {
    throw 'From address cannot be empty - the relay silently drops mail from unauthorized senders.'
}

$secPwd    = Read-Host 'SMTP password' -AsSecureString
$secPwdCfm = Read-Host 'Confirm password' -AsSecureString

# Compare passwords. Convert to plaintext only for the comparison, and always
# free the unmanaged BSTR buffers (ZeroFreeBSTR zeroes them first). Honest
# caveat: PtrToStringBSTR still copies the password into managed strings that
# live until garbage-collected - full scrubbing is not possible in .NET; this
# just minimizes the exposure window.
$bstr1 = [IntPtr]::Zero
$bstr2 = [IntPtr]::Zero
try {
    $bstr1  = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secPwd)
    $bstr2  = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secPwdCfm)
    $plain1 = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr1)
    $plain2 = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr2)
    if ($plain1 -ne $plain2) {
        throw 'Passwords do not match. Re-run New-CredStore.ps1.'
    }
} finally {
    if ($bstr1 -ne [IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr1) }
    if ($bstr2 -ne [IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr2) }
    $plain1 = $null; $plain2 = $null
}

# Generate AES-256 key (32 bytes). RandomNumberGenerator::Create() is the
# non-obsolete factory (RNGCryptoServiceProvider is deprecated) and works on
# both Windows PowerShell 5.1 (.NET Framework) and PowerShell 7+ (.NET).
$rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
$key = New-Object byte[] 32
$rng.GetBytes($key)
$rng.Dispose()

# Encrypt password with AES key
$encrypted = ConvertFrom-SecureString -SecureString $secPwd -Key $key

# Save files
[System.IO.File]::WriteAllBytes($keyFile, $key)
[System.IO.File]::WriteAllText($credFile, $encrypted)
[System.IO.File]::WriteAllText($userFile, $SmtpUser)

# Per-deployment SMTP settings (no secrets here - just connection details).
# ServiceMonitor.ps1 auto-loads this file from next to the cred file. Written
# without a BOM so ConvertFrom-Json is happy on Windows PowerShell 5.1.
$smtpSettings = [ordered]@{
    server = $SmtpServer
    port   = $SmtpPort
    from   = $FromAddress
    user   = $SmtpUser
    useSsl = [bool]$SmtpUseSsl
}
$json = $smtpSettings | ConvertTo-Json
[System.IO.File]::WriteAllText($jsonFile, $json, (New-Object System.Text.UTF8Encoding($false)))

Write-Host ''
Write-Host 'Files written:' -ForegroundColor Cyan
Write-Host "  Key  : $keyFile"
Write-Host "  Cred : $credFile"
Write-Host "  User : $userFile"
Write-Host "  Json : $jsonFile  (server $SmtpServer`:$SmtpPort, ssl: $([bool]$SmtpUseSsl))"

# Restrict ACL on key file: SYSTEM + Administrators only
try {
    $acl = Get-Acl -LiteralPath $keyFile
    $acl.SetAccessRuleProtection($true, $false)    # break inheritance, discard inherited rules

    foreach ($rule in @($acl.Access)) {
        $acl.RemoveAccessRule($rule) | Out-Null
    }

    $system = New-Object Security.Principal.SecurityIdentifier 'S-1-5-18'       # NT AUTHORITY\SYSTEM
    $admins = New-Object Security.Principal.SecurityIdentifier 'S-1-5-32-544'   # BUILTIN\Administrators

    $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule(
        $system, 'FullControl', 'None', 'None', 'Allow')))
    $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule(
        $admins, 'FullControl', 'None', 'None', 'Allow')))

    Set-Acl -LiteralPath $keyFile -AclObject $acl
    Write-Host ''
    Write-Host "ACL restricted on smtp.key (SYSTEM + Administrators only)." -ForegroundColor Green
} catch {
    Write-Warning "Could not restrict ACL on key file: $_ - RESTRICT IT MANUALLY before deploying."
}

# Verify round-trip decryption
try {
    $verifyKey = [System.IO.File]::ReadAllBytes($keyFile)
    $verifyEnc = [System.IO.File]::ReadAllText($credFile).Trim()
    $verifySec = ConvertTo-SecureString -String $verifyEnc -Key $verifyKey
    $verifyCred = New-Object System.Net.NetworkCredential($SmtpUser, $verifySec)
    # (credential object created - not used further)
    $verifyCred = $null
    Write-Host 'Decryption round-trip: OK' -ForegroundColor Green
} catch {
    Write-Warning "Round-trip verification FAILED: $_ - the credential files may be corrupt."
}

Write-Host ''
Write-Host 'Usage with ServiceMonitor.ps1:' -ForegroundColor Cyan
Write-Host '  ServiceMonitor.ps1 auto-loads smtp.json (server/port/from/user/ssl)'
Write-Host '  from next to the cred file, so you only need:'
Write-Host "    -SmtpKeyFile '$keyFile' ``"
Write-Host "    -SmtpCredFile '$credFile'"
Write-Host ''
Write-Host 'Test connectivity before scheduling:'
Write-Host "  .\ServiceMonitor.ps1 -TestEmail -SmtpKeyFile '$keyFile' -SmtpCredFile '$credFile'"
