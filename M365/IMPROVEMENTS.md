# Improvements & Enhancement Proposals

## Get-LicensedUsers.ps1

This document tracks enhancement proposals and reference implementations for `Get-LicensedUsers.ps1`.

### Overview

`Get-LicensedUsers.ps1` exports detailed licensed user information from M365 tenants, including:
- User identities (UPN, Display Name, Department)
- License SKUs and assignment dates
- Last sign-in activity
- MFA status

Output is generated in both HTML and CSV formats.

---

## Key Improvements

### 1. Robust Error Handling with Retry Logic

**Problem:** Microsoft Graph API calls can fail transiently due to throttling or temporary service issues.

**Solution:** Implement retry mechanism with exponential backoff.

**Implementation:**
```powershell
function Invoke-GraphWithRetry {
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 5
    )
    
    $attempt = 0
    $lastError = $null
    
    while ($attempt -lt $MaxRetries) {
        try {
            return & $ScriptBlock
        }
        catch {
            $attempt++
            $lastError = $_
            
            if ($attempt -lt $MaxRetries) {
                $delay = $RetryDelaySeconds * $attempt
                Write-Warning "Attempt $attempt failed: $($_.Exception.Message). Retrying in $delay seconds..."
                Start-Sleep -Seconds $delay
            }
        }
    }
    
    throw "Operation failed after $MaxRetries attempts. Last error: $($lastError.Exception.Message)"
}
```

**Benefit:** Improves reliability for large-scale queries and unreliable network conditions.

---

### 2. Progress Reporting for Large Datasets

**Problem:** Long-running queries on large tenants provide no feedback to the user.

**Solution:** Display real-time progress bar and user count during data retrieval.

**Implementation:**
```powershell
function Get-AllUsersWithProgress {
    param([string]$Message = "Retrieving users")
    
    Write-Progress -Activity $Message -Status "Initializing..." -PercentComplete 0
    
    $allUsers = @()
    $pageSize = 999
    
    try {
        # Get total count first (if available)
        $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users?`$count=true&`$top=1" -Headers @{"ConsistencyLevel"="eventual"}
        $totalCount = $response.'@odata.count'
        
        # Paginate through results
        $uri = "https://graph.microsoft.com/v1.0/users?`$select=Id,AssignedLicenses,DisplayName,UserPrincipalName,Mail&`$top=$pageSize"
        
        while ($uri) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
            
            if ($response.value) {
                $allUsers += $response.value
                
                if ($totalCount) {
                    $percent = [Math]::Min(100, ($allUsers.Count / $totalCount) * 100)
                    Write-Progress -Activity $Message -Status "Retrieved $($allUsers.Count) of $totalCount users" -PercentComplete $percent
                }
                else {
                    Write-Progress -Activity $Message -Status "Retrieved $($allUsers.Count) users" -PercentComplete -1
                }
            }
            
            $uri = $response.'@odata.nextLink'
        }
        
        Write-Progress -Activity $Message -Completed
        return $allUsers
    }
    catch {
        Write-Progress -Activity $Message -Completed
        throw
    }
}
```

**Benefit:** Provides user feedback during long operations; displays estimated completion time.

---

### 3. Input Validation

**Problem:** Invalid parameters or missing CSV files cause cryptic errors mid-execution.

**Solution:** Validate all inputs upfront with clear error messages.

**Implementation:**
```powershell
function Test-CsvOverridePath {
    param([string]$Path)
    
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $true  # Optional parameter
    }
    
    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        throw "CSV override file not found: $Path"
    }
    
    $extension = [System.IO.Path]::GetExtension($Path)
    if ($extension -ne '.csv') {
        throw "Override file must be a CSV file, got: $extension"
    }
    
    # Validate CSV structure
    try {
        $testImport = Import-Csv -Path $Path -ErrorAction Stop
        
        $requiredColumns = @('UPN', 'Email', 'UserPrincipalName')
        $hasIdentifier = $testImport[0].PSObject.Properties.Name | Where-Object { $_ -in $requiredColumns }
        
        if (-not $hasIdentifier) {
            throw "CSV must contain at least one identifier column: UPN, Email, or UserPrincipalName"
        }
        
        return $true
    }
    catch {
        throw "Invalid CSV format: $($_.Exception.Message)"
    }
}
```

**Benefit:** Fail fast with actionable error messages; prevents partial data collection.

---

### 4. Memory Optimization for Large Tenants

**Problem:** Processing large user arrays in memory can exhaust available RAM.

**Solution:** Process users in batches and periodically trigger garbage collection.

**Implementation:**
```powershell
function Get-LicensedUsersOptimized {
    param(
        [array]$Users,
        [hashtable]$SignInLookup,
        [hashtable]$AdminRoleLookup,
        [hashtable]$SkuMap,
        [hashtable]$LicenseNameMap
    )
    
    # Process in batches to avoid memory issues
    $batchSize = 1000
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    
    for ($i = 0; $i -lt $Users.Count; $i += $batchSize) {
        $batch = $Users[$i..[Math]::Min($i + $batchSize - 1, $Users.Count - 1)]
        
        foreach ($user in $batch) {
            $licensedUser = Process-SingleUser -User $user -SignInLookup $SignInLookup -AdminRoleLookup $AdminRoleLookup -SkuMap $SkuMap -LicenseNameMap $LicenseNameMap
            
            if ($licensedUser) {
                $results.Add($licensedUser)
            }
        }
        
        # Force garbage collection periodically for very large datasets
        if ($i % 5000 -eq 0 -and $i -gt 0) {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
    
    return $results.ToArray()
}
```

**Benefit:** Handles tenants with 10,000+ users without memory pressure.

---

### 5. Configuration File Support

**Problem:** Changing parameters requires editing the script or passing long command lines.

**Solution:** Support external JSON configuration files with sensible defaults.

**Implementation:**
```powershell
function Get-ReportConfiguration {
    param([string]$ConfigPath)
    
    $defaultConfig = @{
        InactiveThresholdDays = 90
        Language = 'English'
        MaxRetries = 3
        PageSize = 999
        EnableProgressReporting = $true
        CacheResults = $false
    }
    
    if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
        return $defaultConfig
    }
    
    if (Test-Path $ConfigPath) {
        try {
            $fileConfig = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json -AsHashtable
            
            # Merge with defaults
            foreach ($key in $fileConfig.Keys) {
                if ($defaultConfig.ContainsKey($key)) {
                    $defaultConfig[$key] = $fileConfig[$key]
                }
            }
        }
        catch {
            Write-Warning "Could not load configuration file: $_. Using defaults."
        }
    }
    
    return $defaultConfig
}
```

**Benefit:** Easier configuration management; supports different profiles for different tenants.

---

### 6. Structured Logging

**Problem:** Errors and warnings go only to console; no persistent audit trail.

**Solution:** Log all activities to file with timestamps and severity levels.

**Implementation:**
```powershell
function Write-ReportLog {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level,
        
        [Parameter(Mandatory)]
        [string]$Message,
        
        [string]$LogPath
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Console output
    switch ($Level) {
        'Info'    { Write-Host $logEntry -ForegroundColor Cyan }
        'Warning' { Write-Warning $logEntry }
        'Error'   { Write-Error $logEntry }
    }
    
    # File output
    if ($LogPath) {
        try {
            Add-Content -Path $LogPath -Value $logEntry -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not write to log file: $_"
        }
    }
}
```

**Benefit:** Audit trail for compliance; easier troubleshooting of failed runs.

---

### 7. Data Caching for Development/Testing

**Problem:** Repeated runs fetch the same data, wasting API calls and time.

**Solution:** Cache JSON results locally with configurable expiration.

**Implementation:**
```powershell
function Get-CachedData {
    param(
        [Parameter(Mandatory)]
        [string]$CacheKey,
        
        [Parameter(Mandatory)]
        [scriptblock]$DataProvider,
        
        [string]$CacheDir,
        
        [int]$CacheExpirationHours = 24
    )
    
    if (-not $CacheDir) {
        return & $DataProvider
    }
    
    $cacheFile = Join-Path $CacheDir "$CacheKey.json"
    
    if (Test-Path $cacheFile) {
        $cacheAge = (Get-Date) - (Get-Item $cacheFile).LastWriteTime
        
        if ($cacheAge.TotalHours -lt $CacheExpirationHours) {
            Write-Host "Using cached data for $CacheKey (age: $([Math]::Round($cacheAge.TotalHours, 1)) hours)" -ForegroundColor DarkGray
            $data = Get-Content -Path $cacheFile -Raw | ConvertFrom-Json
            return $data
        }
    }
    
    # Fetch fresh data
    Write-Host "Fetching fresh data for $CacheKey..." -ForegroundColor Cyan
    $data = & $DataProvider
    
    # Cache it
    try {
        if (-not (Test-Path $CacheDir)) {
            New-Item -Path $CacheDir -ItemType Directory -Force | Out-Null
        }
        $data | ConvertTo-Json -Depth 10 | Set-Content -Path $cacheFile -Encoding UTF8
    }
    catch {
        Write-Warning "Could not cache data: $_"
    }
    
    return $data
}
```

**Benefit:** Speeds up development cycles; reduces API quota usage during testing.

---

### 8. Template-Based HTML Generation

**Problem:** HTML generation is hardcoded; difficult to customize styling or layout.

**Solution:** Use template files with placeholder replacement.

**Implementation:**
```powershell
function Get-HtmlTemplate {
    param(
        [hashtable]$Data,
        [string]$TemplatePath
    )
    
    if ($TemplatePath -and (Test-Path $TemplatePath)) {
        $template = Get-Content -Path $TemplatePath -Raw
    }
    else {
        $template = Get-DefaultHtmlTemplate
    }
    
    # Replace placeholders
    foreach ($key in $Data.Keys) {
        $placeholder = "{{$key}}"
        $value = $Data[$key]
        $template = $template -replace [regex]::Escape($placeholder), $value
    }
    
    return $template
}
```

**Benefit:** Enables custom branding and report layouts without code changes.

---

### 9. Multiple Export Formats

**Problem:** Only HTML and CSV are supported; no JSON or Excel output.

**Solution:** Add export format options with Excel support via ImportExcel module.

**Implementation:**
```powershell
function Export-ReportData {
    param(
        [Parameter(Mandatory)]
        [array]$Data,
        
        [Parameter(Mandatory)]
        [ValidateSet('HTML', 'CSV', 'JSON', 'Excel')]
        [string]$Format,
        
        [Parameter(Mandatory)]
        [string]$OutputPath
    )
    
    switch ($Format) {
        'HTML' {
            # Existing HTML generation
        }
        'CSV' {
            $Data | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        }
        'JSON' {
            $Data | ConvertTo-Json -Depth 10 | Set-Content -Path $OutputPath -Encoding UTF8
        }
        'Excel' {
            if (Get-Module -ListAvailable -Name ImportExcel) {
                Import-Module ImportExcel
                $Data | Export-Excel -Path $OutputPath -AutoSize -FreezeTopRow -BoldTopRow
            }
            else {
                Write-Warning "ImportExcel module not available. Install with: Install-Module ImportExcel"
                throw "Excel export requires ImportExcel module"
            }
        }
    }
}
```

**Benefit:** Supports downstream tools that consume JSON or Excel; easier sharing and formatting.

---

### 10. Scheduled Execution via Windows Task Scheduler

**Problem:** No built-in support for automated, recurring reports.

**Solution:** Register a Windows Scheduled Task for daily/weekly/monthly runs.

**Implementation:**
```powershell
function Register-ScheduledReport {
    param(
        [Parameter(Mandatory)]
        [string]$ScriptPath,
        
        [string]$TaskName = "M365-LicensedUsersReport",
        
        [ValidateSet('Daily', 'Weekly', 'Monthly')]
        [string]$Frequency = 'Weekly',
        
        [DateTime]$StartTime = (Get-Date).AddHours(1)
    )
    
    $action = New-ScheduledTaskAction -Execute 'pwsh.exe' -Argument "-File `"$ScriptPath`""
    
    $trigger = switch ($Frequency) {
        'Daily'   { New-ScheduledTaskTrigger -Daily -At $StartTime }
        'Weekly'  { New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At $StartTime }
        'Monthly' { New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At $StartTime }
    }
    
    $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive
    
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
    
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force
    
    Write-Host "Scheduled task '$TaskName' created successfully" -ForegroundColor Green
}
```

**Benefit:** Enables unattended, recurring license audits without manual intervention.

---

## Summary

| Improvement | Benefit | Priority |
|---|---|---|
| Error Handling & Retry | Reliability | High |
| Progress Reporting | User Feedback | Medium |
| Input Validation | Fail Fast | High |
| Memory Optimization | Large Tenant Support | High |
| Configuration Files | Flexibility | Medium |
| Structured Logging | Audit Trail | Medium |
| Data Caching | Development Speed | Low |
| Template HTML | Customization | Medium |
| Multiple Export Formats | Interoperability | Medium |
| Scheduled Execution | Automation | Medium |

---

**Last Updated:** 2026-05-05  
**Source:** Consolidated from Scanning-Itwhatcher  
**Repository:** [Windows-Admin-Tools](https://github.com/ZirHuan/Windows-Admin-Tools)
