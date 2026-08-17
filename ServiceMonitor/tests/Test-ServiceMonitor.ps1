#Requires -Version 5.1
<#
.SYNOPSIS
    Pester v5 integration tests for ServiceMonitor.ps1.

.DESCRIPTION
    Runs functional tests against the real ServiceMonitor.ps1 script using
    temporary files in $env:TEMP. No real SMTP is used (-NoEmail throughout).
    No services are stopped or modified - tests use already-running services
    (Spooler) or fake service names.

    Prerequisites:
        Install-Module -Name Pester -MinimumVersion 5.0 -Force -Scope CurrentUser

    Usage:
        Invoke-Pester .\Test-ServiceMonitor.ps1 -Output Detailed
        Invoke-Pester .\Test-ServiceMonitor.ps1 -Output Detailed -Tag Quick

.NOTES
    Version: 1.2.0
    Requires: Pester 5.0+, Windows PowerShell 5.1 or PowerShell 7+, Run as Admin
              (needed so ServiceMonitor.ps1 can call Get-Service without errors)
#>

BeforeAll {
    $script:SourceScript = Join-Path $PSScriptRoot '..\ServiceMonitor.ps1'
    if (-not (Test-Path -LiteralPath $script:SourceScript)) {
        throw "ServiceMonitor.ps1 not found at: $script:SourceScript"
    }

    # Working directory for all test artefacts
    $script:TestRoot = Join-Path $env:TEMP "SMTest_$(Get-Random)"
    New-Item -ItemType Directory -Path $script:TestRoot | Out-Null

    # Run a COPY of the script from an otherwise-empty directory. Several paths
    # are resolved against the script's own folder ($ScriptDir): monitor-config.json,
    # ServiceMonitor.state.json and smtp.json. Running the repo copy in place would
    # (a) load the repo's real monitor-config.json instead of the legacy flat files
    # these tests exercise, and (b) write state/log artefacts into the working tree.
    # An empty folder gives legacy flat-file mode via the documented default path -
    # passing a nonexistent -ConfigFile is a FATAL error by design (see the
    # 'Explicit missing config file' block below), so it cannot be used here.
    $script:ScriptPath = Join-Path $script:TestRoot 'ServiceMonitor.ps1'
    Copy-Item -LiteralPath $script:SourceScript -Destination $script:ScriptPath -Force

    # Single source of truth for the version assertion in the startup-banner test.
    $versionMatch = [regex]::Match(
        [System.IO.File]::ReadAllText($script:SourceScript),
        "(?m)^\s*\`$ScriptVersion\s*=\s*'([^']+)'")
    if (-not $versionMatch.Success) {
        throw 'Could not parse $ScriptVersion from ServiceMonitor.ps1'
    }
    $script:ScriptVersion = $versionMatch.Groups[1].Value

    $script:Recipients  = Join-Path $script:TestRoot 'test-recipients.txt'
    # Needs at least one real recipient so the alert path (and its [NoEmail]
    # suppression logging) is actually exercised; -NoEmail means nothing is sent.
    @('# test recipients', 'test-noreply@example.com') |
        Set-Content -Path $script:Recipients -Encoding UTF8

    function Invoke-SM {
        # Helper: run ServiceMonitor.ps1 with -NoEmail and a fresh log, return log content.
        # -ConfigFile is deliberately NOT passed: the script copy lives in an empty
        # directory, so the default monitor-config.json is absent and legacy
        # flat-file mode is selected.
        param(
            [string]   $ServicesContent,
            [string[]] $ExtraArgs = @()
        )
        $id         = [System.IO.Path]::GetRandomFileName()
        $svcFile    = Join-Path $script:TestRoot "svc_$id.txt"
        $logFile    = Join-Path $script:TestRoot "log_$id.log"
        [System.IO.File]::WriteAllText($svcFile, $ServicesContent)

        & $script:ScriptPath `
            -ServicesFile   $svcFile `
            -RecipientsFile $script:Recipients `
            -LogFile        $logFile `
            -NoEmail `
            @ExtraArgs 2>&1 | Out-Null

        if (Test-Path -LiteralPath $logFile) {
            return [System.IO.File]::ReadAllText($logFile)
        }
        return ''
    }
}

AfterAll {
    Remove-Item -Path $script:TestRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------------
Describe 'Startup banner' -Tag Quick {
    It 'Logs version and hostname in startup banner' {
        # The expected version is read from the script rather than hardcoded: a
        # literal here silently rotted from v1.2.5 through the v1.3.0 release and
        # only surfaced once an unrelated CI break was fixed.
        $log = Invoke-SM -ServicesContent 'Spooler'
        $log | Should -Match ('ServiceMonitor v' + [regex]::Escape($script:ScriptVersion))
        $log | Should -Match $env:COMPUTERNAME
    }
}

# ---------------------------------------------------------------------------
Describe 'Running service' -Tag Quick {
    It 'Logs OK for Print Spooler (expected to be running)' {
        $log = Invoke-SM -ServicesContent 'Spooler'
        $log | Should -Match "OK: 'Spooler' is running"
    }

    It 'Does not log WARNING or ERROR for a healthy service' {
        $log = Invoke-SM -ServicesContent 'Spooler'
        # Legacy flat-file mode always logs ONE warning that monitor-config.json
        # is absent; that is expected here and unrelated to service health.
        $relevant = ($log -split "`r?`n" |
            Where-Object { $_ -notmatch 'falling back to services.txt' }) -join "`n"
        $relevant | Should -Not -Match '\[WARNING\]'
        $relevant | Should -Not -Match '\[ERROR  \]'
    }

    It 'Exits with 0 when all services are healthy' {
        $id      = [System.IO.Path]::GetRandomFileName()
        $svcFile = Join-Path $script:TestRoot "svc_$id.txt"
        $logFile = Join-Path $script:TestRoot "log_$id.log"
        'Spooler' | Set-Content -Path $svcFile -Encoding UTF8

        & $script:ScriptPath `
            -ServicesFile    $svcFile `
            -RecipientsFile  $script:Recipients `
            -LogFile         $logFile `
            -NoEmail 2>&1 | Out-Null

        $LASTEXITCODE | Should -Be 0
    }
}

# ---------------------------------------------------------------------------
Describe 'Paused services' -Tag Quick {
    It 'Skips a service prefixed with *' {
        $log = Invoke-SM -ServicesContent "* NonExistentService_PAUSED`nSpooler"
        $log | Should -Match 'Paused.*NonExistentService_PAUSED'
        $log | Should -Match "OK: 'Spooler'"
    }

    It 'Skips a service prefixed with -' {
        $log = Invoke-SM -ServicesContent "- NonExistentService_PAUSED`nSpooler"
        $log | Should -Match 'Paused.*NonExistentService_PAUSED'
    }

    It 'Does not attempt to check or restart a paused service' {
        $log = Invoke-SM -ServicesContent '* NonExistentService_PAUSED'
        # 'not found on' is the service-not-found message; plain 'not found' would
        # also match the unrelated "monitor-config.json not found" fallback warning.
        $log | Should -Not -Match 'not found on'
        $log | Should -Not -Match 'Restart attempt'
    }
}

# ---------------------------------------------------------------------------
Describe 'Comment and blank lines in services.txt' -Tag Quick {
    It 'Ignores comment lines and blank lines' {
        $content = @"
# This is a comment
Spooler

# Another comment
"@
        $log = Invoke-SM -ServicesContent $content
        $log | Should -Match "OK: 'Spooler'"
        $log | Should -Not -Match 'comment'
    }
}

# ---------------------------------------------------------------------------
Describe 'Nonexistent service' -Tag Quick {
    It 'Logs ERROR when service is not found on host' {
        $log = Invoke-SM -ServicesContent 'NonExistentService_XYZ_12345'
        $log | Should -Match '\[ERROR  \]'
        $log | Should -Match 'not found on'
    }

    It 'Logs [NoEmail] for missing service when -NoEmail is set' {
        $log = Invoke-SM -ServicesContent 'NonExistentService_XYZ_12345'
        $log | Should -Match '\[NoEmail\]'
    }

    It 'Exits with 2 when a service is not found (permanent failure)' {
        $id      = [System.IO.Path]::GetRandomFileName()
        $svcFile = Join-Path $script:TestRoot "svc_$id.txt"
        $logFile = Join-Path $script:TestRoot "log_$id.log"
        'NonExistentService_XYZ_12345' | Set-Content -Path $svcFile -Encoding UTF8

        & $script:ScriptPath `
            -ServicesFile    $svcFile `
            -RecipientsFile  $script:Recipients `
            -LogFile         $logFile `
            -NoEmail 2>&1 | Out-Null

        # v1.3.0 split the exit codes: 1 is now reserved for a FATAL script error
        # (bad config, unwritable log), while 2 means the run completed and at
        # least one service failed. A missing service is the latter.
        $LASTEXITCODE | Should -Be 2
    }
}

# ---------------------------------------------------------------------------
Describe 'Service name injection prevention' -Tag Quick {
    It 'Rejects a name with semicolons' {
        $log = Invoke-SM -ServicesContent 'badname; rm -rf /'
        $log | Should -Match 'Rejected invalid service name'
    }

    It 'Rejects a name with pipe characters' {
        $log = Invoke-SM -ServicesContent 'svc | whoami'
        $log | Should -Match 'Rejected invalid service name'
    }

    It 'Accepts a name with allowed special characters (dot, hyphen, dollar, space)' {
        # This name won't exist but should pass validation and hit "not found"
        $log = Invoke-SM -ServicesContent 'My.Valid-Service$Name'
        $log | Should -Not -Match 'Rejected invalid service name'
        $log | Should -Match 'not found on'
    }
}

# ---------------------------------------------------------------------------
Describe 'Empty services file' -Tag Quick {
    # v1.3.0 deliberately downgraded "nothing to check" from an error to a benign
    # outcome: a fresh install has zero services until the operator adds them in
    # the web UI, and exiting 1 spammed an Error event every 5 minutes.
    It 'Warns and exits 0 when services file has only comments' {
        $id      = [System.IO.Path]::GetRandomFileName()
        $svcFile = Join-Path $script:TestRoot "svc_$id.txt"
        $logFile = Join-Path $script:TestRoot "log_$id.log"
        '# only comments here' | Set-Content -Path $svcFile -Encoding UTF8

        & $script:ScriptPath `
            -ServicesFile    $svcFile `
            -RecipientsFile  $script:Recipients `
            -LogFile         $logFile `
            -NoEmail 2>&1 | Out-Null

        $log = [System.IO.File]::ReadAllText($logFile)
        $log | Should -Match 'No services configured yet'
        $log | Should -Match '\[WARNING\]'
        $LASTEXITCODE | Should -Be 0
    }

    It 'Logs error when services file does not exist' {
        $id      = [System.IO.Path]::GetRandomFileName()
        $svcFile = Join-Path $script:TestRoot "svc_nonexistent_$id.txt"
        $logFile = Join-Path $script:TestRoot "log_$id.log"

        & $script:ScriptPath `
            -ServicesFile    $svcFile `
            -RecipientsFile  $script:Recipients `
            -LogFile         $logFile `
            -NoEmail 2>&1 | Out-Null

        $log = [System.IO.File]::ReadAllText($logFile)
        # Must be the Services-file error specifically; a bare 'not found' would
        # also match the benign 'monitor-config.json not found' fallback warning
        # that legacy mode always logs (cf. b2c47d9).
        $log | Should -Match 'Services file not found'
    }
}

# ---------------------------------------------------------------------------
Describe 'Log rotation' {
    It 'Rotates log to .1 when LogMaxSizeMB is 0 (any content triggers rotation)' {
        $id      = [System.IO.Path]::GetRandomFileName()
        $svcFile = Join-Path $script:TestRoot "svc_$id.txt"
        $logFile = Join-Path $script:TestRoot "log_$id.log"
        'Spooler' | Set-Content -Path $svcFile -Encoding UTF8

        # Pre-populate log with content so it is non-empty
        'existing log content' | Set-Content -Path $logFile -Encoding UTF8

        & $script:ScriptPath `
            -ServicesFile    $svcFile `
            -RecipientsFile  $script:Recipients `
            -LogFile         $logFile `
            -LogMaxSizeMB    0 `
            -NoEmail 2>&1 | Out-Null

        Test-Path -LiteralPath "$logFile.1" | Should -BeTrue
    }

    It 'Keeps up to 3 backup files (.1 .2 .3)' {
        $id      = [System.IO.Path]::GetRandomFileName()
        $logFile = Join-Path $script:TestRoot "log_rotate_$id.log"
        $svcFile = Join-Path $script:TestRoot "svc_$id.txt"
        'Spooler' | Set-Content -Path $svcFile -Encoding UTF8

        # Simulate existing backups
        'run 1' | Set-Content "$logFile.1" -Encoding UTF8
        'run 2' | Set-Content "$logFile.2" -Encoding UTF8
        'run 3' | Set-Content "$logFile"   -Encoding UTF8

        & $script:ScriptPath `
            -ServicesFile    $svcFile `
            -RecipientsFile  $script:Recipients `
            -LogFile         $logFile `
            -LogMaxSizeMB    0 `
            -NoEmail 2>&1 | Out-Null

        Test-Path "$logFile.1" | Should -BeTrue
        Test-Path "$logFile.2" | Should -BeTrue
        # .1 should now contain what was in the main log ('run 3')
        (Get-Content "$logFile.1" -Raw) | Should -Match 'run 3'
    }
}

# ---------------------------------------------------------------------------
Describe '-TestEmail mode' -Tag Quick {
    It 'Exits 0 and logs TestEmail complete without checking services' {
        $id      = [System.IO.Path]::GetRandomFileName()
        $logFile = Join-Path $script:TestRoot "log_testemail_$id.log"

        & $script:ScriptPath `
            -RecipientsFile  $script:Recipients `
            -LogFile         $logFile `
            -NoEmail `
            -TestEmail 2>&1 | Out-Null

        $LASTEXITCODE | Should -Be 0
        $log = [System.IO.File]::ReadAllText($logFile)
        $log | Should -Match 'TestEmail'
        $log | Should -Match 'TestEmail complete'
        # Should NOT have loaded or checked any services
        $log | Should -Not -Match 'Active services to check'
    }
}

# ---------------------------------------------------------------------------
Describe 'SMTP credential files' -Tag Quick {
    It 'Logs warning when only one of SmtpCredFile/SmtpKeyFile is provided' {
        $id      = [System.IO.Path]::GetRandomFileName()
        $logFile = Join-Path $script:TestRoot "log_cred_$id.log"
        $fakeKey = Join-Path $script:TestRoot "fake_$id.key"
        [System.IO.File]::WriteAllBytes($fakeKey, (New-Object byte[] 32))

        & $script:ScriptPath `
            -ServicesFile    (Join-Path $PSScriptRoot 'test-services.txt') `
            -RecipientsFile  $script:Recipients `
            -LogFile         $logFile `
            -SmtpKeyFile     $fakeKey `
            -NoEmail 2>&1 | Out-Null

        $log = [System.IO.File]::ReadAllText($logFile)
        $log | Should -Match 'SMTP auth disabled'
    }

    It 'Logs warning when SmtpCredFile does not exist' {
        $id      = [System.IO.Path]::GetRandomFileName()
        $logFile = Join-Path $script:TestRoot "log_cred_$id.log"
        $fakeKey = Join-Path $script:TestRoot "fake_$id.key"
        [System.IO.File]::WriteAllBytes($fakeKey, (New-Object byte[] 32))

        & $script:ScriptPath `
            -ServicesFile    (Join-Path $PSScriptRoot 'test-services.txt') `
            -RecipientsFile  $script:Recipients `
            -LogFile         $logFile `
            -SmtpUser        'user@example.com' `
            -SmtpKeyFile     $fakeKey `
            -SmtpCredFile    'C:\does\not\exist.cred' `
            -NoEmail 2>&1 | Out-Null

        $log = [System.IO.File]::ReadAllText($logFile)
        $log | Should -Match 'SMTP auth disabled'
    }
}

# ---------------------------------------------------------------------------
Describe 'Summary line' -Tag Quick {
    It 'Logs check complete summary with failure count' {
        $log = Invoke-SM -ServicesContent "Spooler`nNonExistentService_XYZ_12345"
        $log | Should -Match 'Check complete\. Failures: 1 / 2'
    }

    It 'Logs 0 failures when all services are healthy' {
        $log = Invoke-SM -ServicesContent 'Spooler'
        $log | Should -Match 'Failures: 0 / 1'
    }
}

# ---------------------------------------------------------------------------
Describe 'Explicit missing config file' -Tag Quick {
    # v1.3.0 hardening: an explicitly passed -ConfigFile that does not exist is an
    # operator error (typo in the scheduled-task action). Silently falling back to
    # a possibly-stale services.txt would monitor the wrong set forever, so the
    # script must fail loudly instead.
    It 'Fails fatally instead of falling back to legacy flat files' {
        $id      = [System.IO.Path]::GetRandomFileName()
        $svcFile = Join-Path $script:TestRoot "svc_$id.txt"
        $logFile = Join-Path $script:TestRoot "log_$id.log"
        $missing = Join-Path $script:TestRoot "absent-config-$id.json"   # must not exist
        'Spooler' | Set-Content -Path $svcFile -Encoding UTF8

        & $script:ScriptPath `
            -ConfigFile      $missing `
            -ServicesFile    $svcFile `
            -RecipientsFile  $script:Recipients `
            -LogFile         $logFile `
            -NoEmail 2>&1 | Out-Null

        $LASTEXITCODE | Should -Be 1
        $log = [System.IO.File]::ReadAllText($logFile)
        $log | Should -Match 'Config file explicitly specified but not found'
        # Must NOT have silently used the legacy services file
        $log | Should -Not -Match "OK: 'Spooler'"
    }
}
