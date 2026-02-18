<#
.SYNOPSIS
    Per-Process CPU Power Meter - Multi-Source Edition
.DESCRIPTION
    Monitors CPU-specific power consumption for selected processes with
    support for multiple power measurement backends:
    - LibreHardwareMonitor (RAPL)
    - Intel PCM (Performance Counter Monitor)
    - Windows Power Meter (System-wide)
    
    Requires: Administrator privileges for RAPL/PCM sources
#>

param(
    [int]$MeasurementIntervalSeconds = 2
)

# Color scheme
$script:Colors = @{
    Header = 'Cyan'
    ProcessName = 'Yellow'
    Value = 'Green'
    Warning = 'Red'
    Info = 'White'
}

# Power source configuration
$script:PowerSource = $null
$script:SourceType = $null
$script:Computer = $null
$script:CpuHardware = $null
$script:PCMEndpoint = "http://localhost:9738"
$script:PCMProcess = $null

#region LibreHardwareMonitor Functions

function Initialize-LibreHardwareMonitor {
    <#
    .SYNOPSIS
        Initialize LibreHardwareMonitor library
    #>
    $dllPath = Join-Path $PSScriptRoot "LibreHardwareMonitorLib.dll"
    
    if (-not (Test-Path $dllPath)) {
        Write-Host "[X] LibreHardwareMonitorLib.dll not found" -ForegroundColor $script:Colors.Warning
        return $false
    }
    
    try {
        Add-Type -Path $dllPath -ErrorAction Stop
        
        $script:Computer = New-Object LibreHardwareMonitor.Hardware.Computer
        $script:Computer.IsCpuEnabled = $true
        $script:Computer.Open()
        
        # Find CPU hardware
        foreach ($hardware in $script:Computer.Hardware) {
            if ($hardware.HardwareType -eq [LibreHardwareMonitor.Hardware.HardwareType]::Cpu) {
                $script:CpuHardware = $hardware
                Write-Host "[OK] CPU detected: $($hardware.Name)" -ForegroundColor Green
                break
            }
        }
        
        if ($null -eq $script:CpuHardware) {
            Write-Host "[X] Could not detect CPU hardware" -ForegroundColor $script:Colors.Warning
            return $false
        }
        
        # Test RAPL sensor
        $script:CpuHardware.Update()
        $powerSensor = $script:CpuHardware.Sensors | Where-Object {
            $_.SensorType -eq [LibreHardwareMonitor.Hardware.SensorType]::Power -and 
            $_.Name -like "*Package*"
        } | Select-Object -First 1
        
        if ($null -eq $powerSensor) {
            Write-Host "[X] RAPL sensor not found" -ForegroundColor $script:Colors.Warning
            return $false
        }
        
        Write-Host "[OK] RAPL sensor available" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "[X] Failed to initialize LibreHardwareMonitor: $_" -ForegroundColor $script:Colors.Warning
        return $false
    }
}

function Get-CpuPowerFromLibreHardwareMonitor {
    if ($null -eq $script:CpuHardware) {
        return $null
    }
    
    try {
        $script:CpuHardware.Update()
        
        $powerSensor = $script:CpuHardware.Sensors | Where-Object {
            $_.SensorType -eq [LibreHardwareMonitor.Hardware.SensorType]::Power -and 
            $_.Name -like "*Package*"
        } | Select-Object -First 1
        
        if ($null -ne $powerSensor -and $null -ne $powerSensor.Value) {
            return [double]$powerSensor.Value * 1000
        }
        
        return $null
    }
    catch {
        return $null
    }
}

#endregion

#region Intel PCM Functions

function Download-IntelPCM {
    <#
    .SYNOPSIS
        Download Intel PCM if not present
    #>
    $pcmDir = Join-Path $PSScriptRoot "pcm"
    $pcmExe = Join-Path $pcmDir "pcm-sensor-server.exe"
    
    if (Test-Path $pcmExe) {
        return $pcmDir
    }
    
    Write-Host "`n[!] Intel PCM not found. Would you like to download it? (Y/N): " -NoNewline -ForegroundColor Yellow
    $response = Read-Host
    
    if ($response -ne 'Y' -and $response -ne 'y') {
        return $null
    }
    
    try {
        Write-Host "[*] Downloading Intel PCM..." -ForegroundColor Cyan
        
        $zipPath = Join-Path $PSScriptRoot "pcm-temp.zip"
        $downloadSuccess = $false
        
        # Try AppVeyor first
        try {
            Write-Host "[*] Trying AppVeyor builds..." -ForegroundColor Gray
            $apiUrl = "https://ci.appveyor.com/api/projects/opcm/pcm"
            $project = Invoke-RestMethod -Uri $apiUrl -TimeoutSec 10 -ErrorAction Stop
            
            $buildId = $project.build.buildId
            $jobId = $project.build.jobs[0].jobId
            
            Write-Host "[*] Found build ID: $buildId" -ForegroundColor Gray
            
            $artifactUrl = "https://ci.appveyor.com/api/buildjobs/$jobId/artifacts"
            $artifacts = Invoke-RestMethod -Uri $artifactUrl -TimeoutSec 10 -ErrorAction Stop
            
            # Try different naming patterns
            $windowsArtifact = $artifacts | Where-Object { 
                ($_.fileName -like "*windows*" -or $_.fileName -like "*win*" -or $_.fileName -like "*.zip") -and
                $_.fileName -notlike "*linux*" -and 
                $_.fileName -notlike "*mac*"
            } | Select-Object -First 1
            
            if ($null -ne $windowsArtifact) {
                $downloadUrl = "https://ci.appveyor.com/api/buildjobs/$jobId/artifacts/$($windowsArtifact.fileName)"
                Write-Host "[*] Downloading from AppVeyor: $($windowsArtifact.fileName)" -ForegroundColor Cyan
                Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -TimeoutSec 60 -ErrorAction Stop
                $downloadSuccess = $true
            }
        }
        catch {
            Write-Host "[!] AppVeyor download failed: $_" -ForegroundColor Yellow
        }
        
        # Fallback to GitHub releases
        if (-not $downloadSuccess) {
            try {
                Write-Host "[*] Trying GitHub releases..." -ForegroundColor Gray
                $releasesUrl = "https://api.github.com/repos/intel/pcm/releases/latest"
                $release = Invoke-RestMethod -Uri $releasesUrl -TimeoutSec 10 -ErrorAction Stop
                
                # Find Windows asset
                $windowsAsset = $release.assets | Where-Object { 
                    $_.name -like "*windows*" -or $_.name -like "*win*.zip"
                } | Select-Object -First 1
                
                if ($null -ne $windowsAsset) {
                    Write-Host "[*] Downloading from GitHub: $($windowsAsset.name)" -ForegroundColor Cyan
                    Invoke-WebRequest -Uri $windowsAsset.browser_download_url -OutFile $zipPath -TimeoutSec 120 -ErrorAction Stop
                    $downloadSuccess = $true
                }
                else {
                    Write-Host "[!] No Windows release found on GitHub" -ForegroundColor Yellow
                }
            }
            catch {
                Write-Host "[!] GitHub download failed: $_" -ForegroundColor Yellow
            }
        }
        
        if (-not $downloadSuccess) {
            Write-Host "[X] Could not download PCM from any source" -ForegroundColor Red
            Write-Host "    Manual download:" -ForegroundColor Gray
            Write-Host "    - AppVeyor: https://ci.appveyor.com/project/opcm/pcm/history" -ForegroundColor Gray
            Write-Host "    - GitHub: https://github.com/intel/pcm/releases" -ForegroundColor Gray
            return $null
        }
        
        Write-Host "[*] Extracting files..." -ForegroundColor Cyan
        
        # Create pcm directory
        if (-not (Test-Path $pcmDir)) {
            New-Item -ItemType Directory -Path $pcmDir -Force | Out-Null
        }
        
        # Extract
        Expand-Archive -Path $zipPath -DestinationPath $pcmDir -Force
        
        # Clean up
        Remove-Item $zipPath -Force
        
        # PCM might be in subdirectory, try to find it
        $possibleExe = Get-ChildItem -Path $pcmDir -Filter "pcm-sensor-server.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        
        if ($null -ne $possibleExe) {
            # If found in subdirectory, move files to root
            if ($possibleExe.DirectoryName -ne $pcmDir) {
                Write-Host "[*] Reorganizing files..." -ForegroundColor Gray
                Get-ChildItem -Path $possibleExe.DirectoryName | Move-Item -Destination $pcmDir -Force
            }
            
            # Check for required OpenSSL DLLs
            $cryptoDll = Join-Path $pcmDir "libcrypto-1_1-x64.dll"
            $sslDll = Join-Path $pcmDir "libssl-1_1-x64.dll"
            
            if (-not (Test-Path $cryptoDll) -or -not (Test-Path $sslDll)) {
                Write-Host "[!] OpenSSL DLLs missing, downloading..." -ForegroundColor Yellow
                
                try {
                    # Download OpenSSL 1.1.1 light (required by PCM)
                    $opensslUrl = "https://slproweb.com/download/Win64OpenSSL_Light-1_1_1w.exe"
                    $opensslInstaller = Join-Path $PSScriptRoot "openssl-temp.exe"
                    
                    Write-Host "[*] Downloading OpenSSL 1.1.1..." -ForegroundColor Cyan
                    Invoke-WebRequest -Uri $opensslUrl -OutFile $opensslInstaller -TimeoutSec 60 -ErrorAction Stop
                    
                    # Extract DLLs using 7zip if available, otherwise try manual extraction
                    Write-Host "[*] Extracting OpenSSL DLLs..." -ForegroundColor Cyan
                    
                    # Try to find system OpenSSL installation
                    $systemCrypto = @(
                        "C:\Windows\System32\libcrypto-1_1-x64.dll",
                        "C:\Program Files\OpenSSL-Win64\bin\libcrypto-1_1-x64.dll",
                        "C:\OpenSSL-Win64\bin\libcrypto-1_1-x64.dll"
                    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
                    
                    $systemSsl = @(
                        "C:\Windows\System32\libssl-1_1-x64.dll",
                        "C:\Program Files\OpenSSL-Win64\bin\libssl-1_1-x64.dll",
                        "C:\OpenSSL-Win64\bin\libssl-1_1-x64.dll"
                    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
                    
                    if ($systemCrypto -and $systemSsl) {
                        Write-Host "[*] Found system OpenSSL, copying DLLs..." -ForegroundColor Gray
                        Copy-Item $systemCrypto -Destination $cryptoDll -Force
                        Copy-Item $systemSsl -Destination $sslDll -Force
                    }
                    else {
                        Write-Host "[!] Automatic OpenSSL extraction not implemented" -ForegroundColor Yellow
                        Write-Host "    Please install OpenSSL manually:" -ForegroundColor Gray
                        Write-Host "    1. Download from: https://slproweb.com/products/Win32OpenSSL.html" -ForegroundColor Gray
                        Write-Host "    2. Install Win64 OpenSSL v1.1.1" -ForegroundColor Gray
                        Write-Host ("    3. Copy these DLLs to {0}:" -f $pcmDir) -ForegroundColor Gray
                        Write-Host "       - libcrypto-1_1-x64.dll" -ForegroundColor Gray
                        Write-Host "       - libssl-1_1-x64.dll" -ForegroundColor Gray
                    }
                    
                    if (Test-Path $opensslInstaller) {
                        Remove-Item $opensslInstaller -Force
                    }
                }
                catch {
                    Write-Host "[!] OpenSSL download failed: $_" -ForegroundColor Yellow
                }
            }
            
            Write-Host "[OK] Intel PCM downloaded successfully!" -ForegroundColor Green
            Write-Host "[*] Location: $pcmDir" -ForegroundColor Gray
            
            if (-not (Test-Path $cryptoDll)) {
                Write-Host "[!] Note: OpenSSL DLLs still needed for PCM to run" -ForegroundColor Yellow
            }
            
            return $pcmDir
        }
        else {
            Write-Host "[X] pcm-sensor-server.exe not found after extraction" -ForegroundColor Red
            Write-Host "    Extracted to: $pcmDir" -ForegroundColor Gray
            Write-Host "    Please check the contents and move files manually if needed" -ForegroundColor Gray
            return $null
        }
    }
    catch {
        Write-Host "[X] Failed to download Intel PCM: $_" -ForegroundColor Red
        Write-Host "    Manual download:" -ForegroundColor Gray
        Write-Host "    - AppVeyor: https://ci.appveyor.com/project/opcm/pcm/history" -ForegroundColor Gray
        Write-Host "    - GitHub: https://github.com/intel/pcm/releases" -ForegroundColor Gray
        return $null
    }
}

function Start-IntelPCMServer {
    <#
    .SYNOPSIS
        Start Intel PCM sensor server
    #>
    param([string]$PCMDir)
    
    $pcmExe = Join-Path $PCMDir "pcm-sensor-server.exe"
    
    if (-not (Test-Path $pcmExe)) {
        return $false
    }
    
    # Check if running as admin
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $isAdmin) {
        Write-Host "[X] Administrator privileges required to start PCM" -ForegroundColor Red
        Write-Host "    Please run PowerShell as Administrator, then:" -ForegroundColor Yellow
        Write-Host ("    cd `"{0}`"" -f $PCMDir) -ForegroundColor Gray
        Write-Host "    .\pcm-sensor-server.exe -p 9738" -ForegroundColor Gray
        return $false
    }
    
    # Check for OpenSSL DLLs
    $cryptoDll = Join-Path $PCMDir "libcrypto-1_1-x64.dll"
    $sslDll = Join-Path $PCMDir "libssl-1_1-x64.dll"
    
    if (-not (Test-Path $cryptoDll) -or -not (Test-Path $sslDll)) {
        Write-Host "[X] OpenSSL DLLs missing in PCM directory" -ForegroundColor Red
        Write-Host "    Required files:" -ForegroundColor Gray
        Write-Host "    - libcrypto-1_1-x64.dll" -ForegroundColor Gray
        Write-Host "    - libssl-1_1-x64.dll" -ForegroundColor Gray
        Write-Host "`n    Quick fix: Install OpenSSL 1.1.1 Light" -ForegroundColor Yellow
        Write-Host "    Download: https://slproweb.com/products/Win32OpenSSL.html" -ForegroundColor Cyan
        Write-Host ("    Then copy DLLs from C:\\Program Files\\OpenSSL-Win64\\bin to {0}" -f $PCMDir) -ForegroundColor Gray
        return $false
    }
    
    # Install MSR driver if needed
    $msrSys = Join-Path $PCMDir "MSR.sys"
    if (Test-Path $msrSys) {
        Write-Host "[*] Checking MSR driver..." -ForegroundColor Cyan
        
        $serviceExists = Get-Service -Name "pcm-msr" -ErrorAction SilentlyContinue
        if (-not $serviceExists) {
            Write-Host "[*] Installing MSR driver service..." -ForegroundColor Cyan
            try {
                # Use sc.exe to create the service
                $result = & sc.exe create "pcm-msr" binPath= $msrSys type= kernel start= demand error= normal 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "[OK] MSR driver service installed" -ForegroundColor Green
                }
                else {
                    Write-Host "[!] MSR driver installation may have failed: $result" -ForegroundColor Yellow
                }
            }
            catch {
                Write-Host "[!] Failed to install MSR driver: $_" -ForegroundColor Yellow
            }
        }
        
        # Start the service
        try {
            $service = Get-Service -Name "pcm-msr" -ErrorAction SilentlyContinue
            if ($service -and $service.Status -ne 'Running') {
                Write-Host "[*] Starting MSR driver service..." -ForegroundColor Cyan
                Start-Service -Name "pcm-msr" -ErrorAction Stop
                Write-Host "[OK] MSR driver service started" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "[!] Could not start MSR service: $_" -ForegroundColor Yellow
        }
    }
    
    Write-Host "[*] Starting Intel PCM sensor server..." -ForegroundColor Cyan
    
    try {
        # Start PCM in background with working directory set
        $processInfo = Start-Process -FilePath $pcmExe -ArgumentList "-p", "9738" -WorkingDirectory $PCMDir -WindowStyle Hidden -PassThru -ErrorAction Stop
        
        # Wait a moment for server to start
        Start-Sleep -Seconds 3
        
        # Check if it's running
        if (-not $processInfo.HasExited) {
            Write-Host "[OK] PCM server started (PID: $($processInfo.Id))" -ForegroundColor Green
            $script:PCMProcess = $processInfo
            return $true
        }
        else {
            Write-Host "[X] PCM server exited immediately (Exit code: $($processInfo.ExitCode))" -ForegroundColor Red
            Write-Host "    Possible issues:" -ForegroundColor Gray
            Write-Host "    - MSR driver not loaded (run pcm.exe once to initialize)" -ForegroundColor Gray
            Write-Host "    - Incompatible CPU" -ForegroundColor Gray
            Write-Host "`n    Try running this first to initialize drivers:" -ForegroundColor Yellow
            Write-Host ("    cd `"{0}`"" -f $PCMDir) -ForegroundColor Gray
            Write-Host "    .\pcm.exe 1 -r" -ForegroundColor Gray
            return $false
        }
    }
    catch {
        Write-Host "[X] Failed to start PCM server: $_" -ForegroundColor Red
        return $false
    }
}

function Initialize-IntelPCM {
    <#
    .SYNOPSIS
        Initialize Intel PCM connection
    #>
    try {
        $response = Invoke-RestMethod -Uri "$script:PCMEndpoint/" -TimeoutSec 2 -ErrorAction Stop
        
        if ($response) {
            Write-Host "[OK] Intel PCM sensor server detected" -ForegroundColor Green
            
            # Test if we can get power metrics
            $testPower = Get-CpuPowerFromPCM
            if ($null -ne $testPower) {
                Write-Host "[OK] PCM power metrics available" -ForegroundColor Green
                return $true
            }
        }
        
        Write-Host "[X] PCM running but no power metrics found" -ForegroundColor $script:Colors.Warning
        return $false
    }
    catch {
        Write-Host "[X] Intel PCM sensor server not available on port 9738" -ForegroundColor $script:Colors.Warning
        
        # Try to download and start PCM
        $pcmDir = Download-IntelPCM
        
        if ($null -ne $pcmDir) {
            $started = Start-IntelPCMServer -PCMDir $pcmDir
            
            if ($started) {
                # Try connecting again
                Start-Sleep -Seconds 2
                try {
                    $response = Invoke-RestMethod -Uri "$script:PCMEndpoint/" -TimeoutSec 2 -ErrorAction Stop
                    if ($response) {
                        Write-Host "[OK] Successfully connected to PCM!" -ForegroundColor Green
                        return $true
                    }
                }
                catch {
                    Write-Host "[X] PCM started but not responding" -ForegroundColor Red
                }
            }
        }
        
        # Show manual instructions
        if ($IsWindows -or $env:OS -match 'Windows') {
            Write-Host ("    Manual: Run PowerShell as Admin, then: .\\pcm\\pcm-sensor-server.exe -p 9738" -f $pcmDir) -ForegroundColor $script:Colors.Info
        }
        else {
            Write-Host "    Manual: sudo ./pcm-sensor-server -p 9738" -ForegroundColor $script:Colors.Info
        }
        return $false
    }
}

function Get-CpuPowerFromPCM {
    try {
        $response = Invoke-RestMethod -Uri "$script:PCMEndpoint/" -TimeoutSec 2 -ErrorAction Stop
        
        # Parse PCM response for CPU package power
        # PCM returns metrics in various formats, looking for package power in Watts
        if ($response -is [string]) {
            # Prometheus format - parse text
            $lines = $response -split "`n"
            foreach ($line in $lines) {
                if ($line -match 'package_power_watts{socket="0"}' -or 
                    $line -match 'cpu_energy_joules{socket="0"}') {
                    $parts = $line -split ' '
                    if ($parts.Count -ge 2) {
                        $watts = [double]$parts[1]
                        return $watts * 1000  # Convert to milliwatts
                    }
                }
            }
        }
        elseif ($response -is [PSCustomObject]) {
            # JSON format
            if ($response.PSObject.Properties['package_power']) {
                return [double]$response.package_power * 1000
            }
        }
        
        return $null
    }
    catch {
        return $null
    }
}

#endregion

#region Power Meter Functions

function Initialize-PowerMeter {
    <#
    .SYNOPSIS
        Initialize Windows Power Meter
    #>
    try {
        $counter = Get-Counter -Counter "\Power Meter(_total)\Power" -ErrorAction Stop
        if ($null -ne $counter) {
            Write-Host "[OK] Windows Power Meter available" -ForegroundColor Green
            return $true
        }
        return $false
    }
    catch {
        Write-Host "[X] Windows Power Meter not available" -ForegroundColor $script:Colors.Warning
        return $false
    }
}

function Get-SystemPowerFromMeter {
    try {
        $counter = Get-Counter -Counter "\Power Meter(_total)\Power" -ErrorAction SilentlyContinue
        if ($null -ne $counter) {
            return [double]$counter.CounterSamples[0].CookedValue
        }
        return $null
    }
    catch {
        return $null
    }
}

#endregion

#region Common Functions

function Get-CpuPowerConsumption {
    <#
    .SYNOPSIS
        Gets current CPU power based on selected source
    #>
    switch ($script:SourceType) {
        "LibreHardwareMonitor" {
            return Get-CpuPowerFromLibreHardwareMonitor
        }
        "IntelPCM" {
            return Get-CpuPowerFromPCM
        }
        "PowerMeter" {
            return Get-SystemPowerFromMeter
        }
        default {
            return $null
        }
    }
}

function Get-SystemPowerConsumption {
    <#
    .SYNOPSIS
        Gets system-wide power (optional, for comparison)
    #>
    try {
        $counter = Get-Counter -Counter "\Power Meter(_total)\Power" -ErrorAction SilentlyContinue
        if ($null -ne $counter) {
            return [double]$counter.CounterSamples[0].CookedValue
        }
        return $null
    }
    catch {
        return $null
    }
}

function Get-ProcessCpuUtilization {
    <#
    .SYNOPSIS
        Gets CPU utilization for all processes
    #>
    try {
        $cpuCounters = Get-Counter -Counter "\Process(*)\% Processor Time" -ErrorAction Stop
        
        $processData = @{}
        $totalCpu = 0
        
        foreach ($sample in $cpuCounters.CounterSamples) {
            $processName = $sample.InstanceName
            $cpuValue = $sample.CookedValue
            
            if ($processName -eq '_total' -or $processName -eq 'idle') {
                continue
            }
            
            if ($processData.ContainsKey($processName)) {
                $processData[$processName] += $cpuValue
            }
            else {
                $processData[$processName] = $cpuValue
            }
            
            $totalCpu += $cpuValue
        }
        
        return @{
            ProcessData = $processData
            TotalCpu = $totalCpu
        }
    }
    catch {
        Write-Warning "Error reading CPU counters: $_"
        return $null
    }
}

function Format-EnergyValue {
    param([double]$Millijoules)
    
    if ($Millijoules -lt 1000) {
        return "{0:N2} mJ" -f $Millijoules
    }
    elseif ($Millijoules -lt 1000000) {
        return "{0:N2} J" -f ($Millijoules / 1000)
    }
    else {
        return "{0:N2} kJ" -f ($Millijoules / 1000000)
    }
}

function Show-PowerSourceMenu {
    <#
    .SYNOPSIS
        Display menu to select power measurement source
    #>
    Clear-Host
    Write-Host "`n=== Power Measurement Source Selection ===" -ForegroundColor $script:Colors.Header
    Write-Host "=" * 60 -ForegroundColor $script:Colors.Header
    Write-Host "`nDetecting available power sources...`n" -ForegroundColor $script:Colors.Info
    
    $sources = @{}
    $index = 1
    
    # Check LibreHardwareMonitor
    $lhmAvailable = Initialize-LibreHardwareMonitor
    if ($lhmAvailable) {
        $sources[$index] = @{
            Name = "LibreHardwareMonitor (RAPL)"
            Type = "LibreHardwareMonitor"
            Description = "CPU package power via Intel RAPL - Most accurate for CPU-only"
        }
        $index++
    }
    
    # Check Intel PCM
    $pcmAvailable = Initialize-IntelPCM
    if ($pcmAvailable) {
        $sources[$index] = @{
            Name = "Intel PCM (Performance Counter Monitor)"
            Type = "IntelPCM"
            Description = "CPU/Memory/System power via PCM - Comprehensive metrics"
        }
        $index++
    }
    
    # Check Power Meter
    $pmAvailable = Initialize-PowerMeter
    if ($pmAvailable) {
        $sources[$index] = @{
            Name = "Windows Power Meter"
            Type = "PowerMeter"
            Description = "System-wide power - Includes all components"
        }
        $index++
    }
    
    if ($sources.Count -eq 0) {
        Write-Host "`nERROR: No power measurement sources available!" -ForegroundColor $script:Colors.Warning
        Write-Host "`nOptions to fix:" -ForegroundColor $script:Colors.Info
        Write-Host "1. Install LibreHardwareMonitorLib.dll for RAPL" -ForegroundColor $script:Colors.Info
        Write-Host "2. Start Intel PCM: pcm-sensor-server" -ForegroundColor $script:Colors.Info
        Write-Host "3. Ensure Power Meter counters are available" -ForegroundColor $script:Colors.Info
        return $null
    }
    
    Write-Host "`nAvailable sources:`n" -ForegroundColor $script:Colors.Info
    
    foreach ($key in $sources.Keys | Sort-Object) {
        $source = $sources[$key]
        Write-Host ("{0}. {1}" -f $key, $source.Name) -ForegroundColor $script:Colors.ProcessName
        Write-Host ("   {0}" -f $source.Description) -ForegroundColor Gray
        Write-Host ""
    }
    
    Write-Host "0. Exit`n" -ForegroundColor $script:Colors.Warning
    Write-Host "Select power source: " -NoNewline -ForegroundColor $script:Colors.Info
    $selection = Read-Host
    
    $selectionNum = 0
    if ([int]::TryParse($selection, [ref]$selectionNum) -and $sources.ContainsKey($selectionNum)) {
        $script:SourceType = $sources[$selectionNum].Type
        Write-Host "`n[OK] Selected: $($sources[$selectionNum].Name)" -ForegroundColor Green
        Start-Sleep -Seconds 1
        return $sources[$selectionNum]
    }
    
    return $null
}

function Show-ProcessList {
    param([hashtable]$ProcessCpuData)
    
    Clear-Host
    Write-Host "`n=== Per-Process CPU Power Meter ===" -ForegroundColor $script:Colors.Header
    Write-Host "Power Source: $script:SourceType" -ForegroundColor Green
    Write-Host "==================================" -ForegroundColor $script:Colors.Header
    Write-Host "`nSelect a process to monitor:`n" -ForegroundColor $script:Colors.Info
    
    $sortedProcesses = $ProcessCpuData.GetEnumerator() | 
        Where-Object { $_.Value -gt 0 } |
        Sort-Object -Property Value -Descending |
        Select-Object -First 30
    
    $index = 1
    $processMap = @{}
    
    foreach ($proc in $sortedProcesses) {
        $processMap[$index] = $proc.Key
        $cpuDisplay = "{0:N2}%" -f $proc.Value
        Write-Host ("{0,3}. {1,-30} CPU: {2}" -f $index, $proc.Key, $cpuDisplay) -ForegroundColor $script:Colors.ProcessName
        $index++
    }
    
    Write-Host "`n  0. Back to source selection" -ForegroundColor $script:Colors.Warning
    Write-Host "  Q. Quit`n" -ForegroundColor $script:Colors.Warning
    
    return $processMap
}

function Start-ProcessMonitoring {
    param(
        [string]$ProcessName,
        [int]$IntervalSeconds
    )
    
    $totalEnergyMillijoules = 0
    $totalEnergyMillijoulesSystem = 0
    $measurementCount = 0
    $startTime = Get-Date
    
    Write-Host "`nStarting monitoring for: $ProcessName" -ForegroundColor $script:Colors.Info
    Write-Host "Press 'Q' to stop...`n" -ForegroundColor $script:Colors.Warning
    Start-Sleep -Seconds 1
    
    while ($true) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq 'Q') {
                break
            }
        }
        
        $cpuData = Get-ProcessCpuUtilization
        $cpuPowerMilliwatts = Get-CpuPowerConsumption
        $systemPowerMilliwatts = Get-SystemPowerConsumption
        
        if ($null -eq $cpuData -or $null -eq $cpuPowerMilliwatts) {
            Write-Host "Error collecting data. Retrying..." -ForegroundColor $script:Colors.Warning
            Start-Sleep -Seconds 1
            continue
        }
        
        $processCpu = 0
        if ($cpuData.ProcessData.ContainsKey($ProcessName)) {
            $processCpu = $cpuData.ProcessData[$ProcessName]
        }
        
        $totalCpu = $cpuData.TotalCpu
        
        if ($totalCpu -gt 0) {
            $processUtilizationRatio = $processCpu / $totalCpu
        }
        else {
            $processUtilizationRatio = 0
        }
        
        # Calculate power allocation
        $processPowerMilliwatts = $cpuPowerMilliwatts * $processUtilizationRatio
        $energyThisPeriodMillijoules = $processPowerMilliwatts * $IntervalSeconds
        $totalEnergyMillijoules += $energyThisPeriodMillijoules
        
        # System power (if available)
        $processPowerMilliwattsSystem = 0
        if ($null -ne $systemPowerMilliwatts -and $script:SourceType -ne "PowerMeter") {
            $processPowerMilliwattsSystem = $systemPowerMilliwatts * $processUtilizationRatio
            $energyThisPeriodMillijoulesSystem = $processPowerMilliwattsSystem * $IntervalSeconds
            $totalEnergyMillijoulesSystem += $energyThisPeriodMillijoulesSystem
        }
        
        $measurementCount++
        $elapsed = (Get-Date) - $startTime
        
        # Display
        Clear-Host
        Write-Host "`n=== Monitoring Process: $ProcessName ===" -ForegroundColor $script:Colors.Header
        Write-Host ("=" * 60) -ForegroundColor $script:Colors.Header
        Write-Host "Power Source: $script:SourceType" -ForegroundColor Green
        Write-Host "Press 'Q' to stop monitoring`n" -ForegroundColor $script:Colors.Warning
        
        Write-Host "Current Measurements:" -ForegroundColor $script:Colors.Info
        
        $sourceLabel = switch ($script:SourceType) {
            "LibreHardwareMonitor" { "CPU Package Power (RAPL)" }
            "IntelPCM" { "CPU Power (PCM)" }
            "PowerMeter" { "System Total Power" }
            default { "Power" }
        }
        
        Write-Host ("  {0}: {1:N2} W" -f $sourceLabel, ($cpuPowerMilliwatts / 1000)) -ForegroundColor $script:Colors.Value
        
        if ($null -ne $systemPowerMilliwatts -and $script:SourceType -ne "PowerMeter") {
            Write-Host ("  System Total Power:       {0:N2} W" -f ($systemPowerMilliwatts / 1000)) -ForegroundColor $script:Colors.Value
        }
        
        Write-Host ("  Process CPU Usage:        {0:N2}%" -f $processCpu) -ForegroundColor $script:Colors.Value
        Write-Host ("  Total CPU Usage:          {0:N2}%" -f $totalCpu) -ForegroundColor $script:Colors.Value
        Write-Host ("  Process Power:            {0:N2} W" -f ($processPowerMilliwatts / 1000)) -ForegroundColor $script:Colors.Value
        
        Write-Host "`nAccumulated Statistics:" -ForegroundColor $script:Colors.Info
        Write-Host ("  Total Energy Consumed:    {0}" -f (Format-EnergyValue $totalEnergyMillijoules)) -ForegroundColor $script:Colors.Value
        Write-Host ("  Average Power:            {0:N2} W" -f (($totalEnergyMillijoules / 1000) / $elapsed.TotalSeconds)) -ForegroundColor $script:Colors.Value
        
        if ($totalEnergyMillijoulesSystem -gt 0) {
            Write-Host "`nSystem Total (for comparison):" -ForegroundColor $script:Colors.Info
            Write-Host ("  Total Energy:             {0}" -f (Format-EnergyValue $totalEnergyMillijoulesSystem)) -ForegroundColor $script:Colors.Value
            Write-Host ("  Average Power:            {0:N2} W" -f (($totalEnergyMillijoulesSystem / 1000) / $elapsed.TotalSeconds)) -ForegroundColor $script:Colors.Value
        }
        
        Write-Host "`nMonitoring Statistics:" -ForegroundColor $script:Colors.Info
        Write-Host ("  Duration:                 {0:hh\:mm\:ss}" -f $elapsed) -ForegroundColor $script:Colors.Value
        Write-Host ("  Measurements:             {0}" -f $measurementCount) -ForegroundColor $script:Colors.Value
        
        Start-Sleep -Seconds $IntervalSeconds
    }
    
    # Final summary
    Write-Host "`n`n=== Monitoring Summary ===" -ForegroundColor $script:Colors.Header
    Write-Host ("Process:              {0}" -f $ProcessName) -ForegroundColor $script:Colors.ProcessName
    Write-Host ("Source:               {0}" -f $script:SourceType) -ForegroundColor $script:Colors.Info
    Write-Host ("Total Energy:         {0}" -f (Format-EnergyValue $totalEnergyMillijoules)) -ForegroundColor $script:Colors.Value
    Write-Host ("Average Power:        {0:N2} W" -f (($totalEnergyMillijoules / 1000) / $elapsed.TotalSeconds)) -ForegroundColor $script:Colors.Value
    Write-Host ("Duration:             {0:hh\:mm\:ss}" -f $elapsed) -ForegroundColor $script:Colors.Value
    Write-Host "`nPress any key to continue..." -ForegroundColor $script:Colors.Info
    $null = [Console]::ReadKey($true)
}

#endregion

#region Main Application

function Start-PowerMeterApp {
    param([int]$IntervalSeconds)
    
    Write-Host "`nPer-Process CPU Power Meter - Multi-Source Edition" -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    
    while ($true) {
        # Select power source
        $selectedSource = Show-PowerSourceMenu
        
        if ($null -eq $selectedSource) {
            Write-Host "`nExiting..." -ForegroundColor $script:Colors.Info
            break
        }
        
        Write-Host "`nPress any key to continue..." -ForegroundColor $script:Colors.Info
        $null = [Console]::ReadKey($true)
        
        # Process monitoring loop
        while ($true) {
            $cpuData = Get-ProcessCpuUtilization
            
            if ($null -eq $cpuData) {
                Write-Host "Error reading CPU data." -ForegroundColor $script:Colors.Warning
                break
            }
            
            $processMap = Show-ProcessList -ProcessCpuData $cpuData.ProcessData
            
            Write-Host "Enter selection: " -NoNewline -ForegroundColor $script:Colors.Info
            $selection = Read-Host
            
            if ($selection -eq 'Q' -or $selection -eq 'q') {
                Write-Host "`nExiting..." -ForegroundColor $script:Colors.Info
                if ($null -ne $script:Computer) {
                    $script:Computer.Close()
                }
                return
            }
            
            if ($selection -eq '0') {
                # Back to source selection
                break
            }
            
            $selectionNum = 0
            if ([int]::TryParse($selection, [ref]$selectionNum) -and $processMap.ContainsKey($selectionNum)) {
                $selectedProcess = $processMap[$selectionNum]
                Start-ProcessMonitoring -ProcessName $selectedProcess -IntervalSeconds $IntervalSeconds
            }
            else {
                Write-Host "`nInvalid selection. Press any key to continue..." -ForegroundColor $script:Colors.Warning
                $null = [Console]::ReadKey($true)
            }
        }
    }
    
    # Cleanup
    if ($null -ne $script:Computer) {
        $script:Computer.Close()
    }
    
    # Stop PCM if we started it
    if ($null -ne $script:PCMProcess -and -not $script:PCMProcess.HasExited) {
        Write-Host "`n[*] Stopping Intel PCM server..." -ForegroundColor Cyan
        $script:PCMProcess.Kill()
        $script:PCMProcess.WaitForExit(5000)
        Write-Host "[OK] PCM server stopped" -ForegroundColor Green
    }
}

# Check admin privileges (required for RAPL/PCM)
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "`nWARNING: Not running as Administrator" -ForegroundColor Yellow
    Write-Host "Some power sources (RAPL, PCM) may not be available.`n" -ForegroundColor Yellow
    Write-Host "Press any key to continue anyway..." -ForegroundColor $script:Colors.Info
    $null = [Console]::ReadKey($true)
}

# Start the application
try {
    Start-PowerMeterApp -IntervalSeconds $MeasurementIntervalSeconds
}
finally {
    # Ensure cleanup happens
    if ($null -ne $script:Computer) {
        $script:Computer.Close()
    }
    if ($null -ne $script:PCMProcess -and -not $script:PCMProcess.HasExited) {
        $script:PCMProcess.Kill()
    }
}

#endregion
