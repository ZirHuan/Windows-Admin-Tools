#Requires -Version 5.1

<#
.SYNOPSIS
    Creates an encrypted SMTP credential store for ServiceMonitor.ps1.

.DESCRIPTION
    Generates a random AES-256 key, prompts for the SMTP password,
    encrypts it using ConvertFrom-SecureString with the AES key, and saves:

        smtp.key   - the AES-256 key (32 bytes, binary)
        smtp.cred  - the AES-encrypted password (base64 text)
        smtp.user  - the SMTP username (plaintext - usernames are not secret)

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
    Version: 1.1.0
#>

[CmdletBinding()]
param(
    [string] $OutputFolder = $PSScriptRoot,
    [string] $SmtpUser     = '',
    [switch] $Force
)

$ErrorActionPreference = 'Stop'

$keyFile  = Join-Path $OutputFolder 'smtp.key'
$credFile = Join-Path $OutputFolder 'smtp.cred'
$userFile = Join-Path $OutputFolder 'smtp.user'

# Guard against accidental overwrite
if (-not $Force) {
    foreach ($f in @($keyFile, $credFile, $userFile)) {
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

$secPwd    = Read-Host 'SMTP password' -AsSecureString
$secPwdCfm = Read-Host 'Confirm password' -AsSecureString

# Compare passwords. Convert to plaintext only for the comparison, and always
# free the unmanaged BSTR buffers (ZeroFreeBSTR zeroes them first) so the
# plaintext password is not left lingering in process memory.
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

Write-Host ''
Write-Host 'Files written:' -ForegroundColor Cyan
Write-Host "  Key  : $keyFile"
Write-Host "  Cred : $credFile"
Write-Host "  User : $userFile"

# Restrict ACL on key file: SYSTEM + Administrators only
try {
    $acl = Get-Acl -Path $keyFile
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

    Set-Acl -Path $keyFile -AclObject $acl
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
Write-Host "  -SmtpUser '$SmtpUser' ``"
Write-Host "  -SmtpKeyFile '$keyFile' ``"
Write-Host "  -SmtpCredFile '$credFile'"
Write-Host ''
Write-Host 'Test connectivity before scheduling:'
Write-Host "  .\ServiceMonitor.ps1 -TestEmail -SmtpUser '$SmtpUser' -SmtpKeyFile '$keyFile' -SmtpCredFile '$credFile'"
