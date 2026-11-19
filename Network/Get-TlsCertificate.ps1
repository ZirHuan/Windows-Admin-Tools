#Requires -Version 5.1
<#
.SYNOPSIS
    Completes a TLS handshake with a host:port and returns the server certificate,
    with expiry in days. Works for HTTPS, LDAPS (636), RDP-TLS, SMTPS (465) and so on.

.DESCRIPTION
    Read-only. Opens a TCP connection, performs a TLS client handshake (certificate
    validation is NOT enforced, so self-signed / expired / wrong-name certs are still
    returned) and reports the certificate the server presented.

    A listening port is not proof the TLS service works - a DC listens on 636 even with
    no usable certificate. Only a completed handshake is, which is what this checks.

    Used on its own, and as a building block by:
      DomainScanner/Invoke-DomainHealthAudit.ps1   (CERT section: LDAPS on every DC)

    Note: implicit TLS only. STARTTLS protocols (SMTP 25/587, LDAP 389 StartTLS) are not
    handled here - use Mail/smtp-relay-tester.ps1 for SMTP.

.PARAMETER ComputerName
    One or more host names or IP addresses. Accepts pipeline input.

.PARAMETER Port
    TCP port. Default 443.

.PARAMETER ServerName
    Name sent as SNI. Default: the ComputerName value.

.PARAMETER TimeoutMs
    Connect and handshake timeout in milliseconds. Default 5000.

.PARAMETER WarnDays
    Status becomes WARN when the certificate expires within this many days. Default 30.

.OUTPUTS
    PSCustomObject: ComputerName, Port, Status (OK/WARN/EXPIRED/ERROR), Subject, DnsName
    (SAN/CN), Issuer, NotBefore, NotAfter, DaysLeft, Thumbprint, Protocol, Error,
    Certificate (X509Certificate2).

.EXAMPLE
    .\Get-TlsCertificate.ps1 -ComputerName www.example.com

.EXAMPLE
    .\Get-TlsCertificate.ps1 -ComputerName dc01.contoso.local,dc02.contoso.local -Port 636 -WarnDays 60 |
        Format-Table ComputerName,Status,DaysLeft,DnsName,Issuer

.NOTES
    Version : 1.0.0
    Read-only.

    Changelog:
      1.0.0  2026-11-19  First version. Extracted from Invoke-DomainHealthAudit
                         (Get-LdapsCertificate) and generalised to any port.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
    [Alias('HostName', 'Target')]
    [string[]]$ComputerName,

    [ValidateRange(1, 65535)]
    [int]$Port = 443,

    [string]$ServerName,

    [ValidateRange(500, 120000)]
    [int]$TimeoutMs = 5000,

    [ValidateRange(0, 3650)]
    [int]$WarnDays = 30
)

begin {
    $ScriptVersion = '1.0.0'
    Write-Verbose "Get-TlsCertificate v$ScriptVersion"
    $acceptAll = [System.Net.Security.RemoteCertificateValidationCallback]{ param($s, $c, $ch, $e) $true }
    # Offer every protocol the runtime knows, so old servers (TLS 1.0) still answer.
    $protocols = [System.Security.Authentication.SslProtocols]::Tls12
    foreach ($n in 'Tls11', 'Tls', 'Tls13') {
        try { $protocols = $protocols -bor [System.Security.Authentication.SslProtocols]::$n } catch { }
    }
}

process {
    foreach ($h in $ComputerName) {
        if ([string]::IsNullOrWhiteSpace($h)) { continue }
        $sni = if ($ServerName) { $ServerName } else { $h }
        $out = [ordered]@{
            ComputerName = $h; Port = $Port; Status = 'ERROR'
            Subject = ''; DnsName = ''; Issuer = ''; NotBefore = $null; NotAfter = $null
            DaysLeft = $null; Thumbprint = ''; Protocol = ''; Error = ''; Certificate = $null
        }
        $tcp = New-Object System.Net.Sockets.TcpClient
        $ssl = $null
        try {
            $ar = $tcp.BeginConnect($h, $Port, $null, $null)
            if (-not $ar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) { throw "TCP $Port timed out after $TimeoutMs ms" }
            $tcp.EndConnect($ar)
            $tcp.ReceiveTimeout = $TimeoutMs
            $tcp.SendTimeout    = $TimeoutMs
            $ssl = New-Object System.Net.Security.SslStream($tcp.GetStream(), $false, $acceptAll)
            $ssl.ReadTimeout  = $TimeoutMs
            $ssl.WriteTimeout = $TimeoutMs
            $ssl.AuthenticateAsClient($sni, $null, $protocols, $false)
            if (-not $ssl.RemoteCertificate) { throw 'Handshake completed but the server sent no certificate' }

            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($ssl.RemoteCertificate)
            $days = [int][math]::Floor(($cert.NotAfter - (Get-Date)).TotalDays)
            $out.Subject     = $cert.Subject
            # DC certs from the Kerberos Authentication template have an empty Subject (SAN only).
            $out.DnsName     = $cert.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::DnsName, $false)
            $out.Issuer      = $cert.Issuer
            $out.NotBefore   = $cert.NotBefore
            $out.NotAfter    = $cert.NotAfter
            $out.DaysLeft    = $days
            $out.Thumbprint  = $cert.Thumbprint
            $out.Protocol    = "$($ssl.SslProtocol)"
            $out.Certificate = $cert
            $out.Status      = if ($days -lt 0) { 'EXPIRED' } elseif ($days -lt $WarnDays) { 'WARN' } else { 'OK' }
        } catch {
            $msg = $_.Exception.Message
            $inner = $_.Exception.GetBaseException().Message
            $out.Error = if ($inner -and $inner -ne $msg) { "$msg ($inner)" } else { $msg }
        } finally {
            if ($ssl) { $ssl.Dispose() }
            $tcp.Close()
        }
        [PSCustomObject]$out
    }
}
