<#
.SYNOPSIS
    Per-Process CPU Power Meter - LibreHardwareMonitor Edition
.DESCRIPTION
    Monitors CPU-specific power consumption for selected processes using
    LibreHardwareMonitor's RAPL interface.
    
    Requires: Administrator privileges and LibreHardwareMonitorLib.dll
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

$script:Computer = $null
$script:CpuHardware = $null

function Initialize-LibreHardwareMonitor {
    <#
    .SYNOPSIS
        Initialize LibreHardwareMonitor library
    #>
    $dllPath = Join-Path $PSScriptRoot "LibreHardwareMonitorLib.dll"
    
    if (-not (Test-Path $dllPath)) {
        Write-Host "ERROR: LibreHardwareMonitorLib.dll not found!" -ForegroundColor $script:Colors.Warning
        Write-Host "Please download from: https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases" -ForegroundColor $script:Colors.Info
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
                Write-Host "CPU detected: $($hardware.Name)" -ForegroundColor Green
                break
            }
        }
        
        if ($null -eq $script:CpuHardware) {
            Write-Host "ERROR: Could not detect CPU hardware" -ForegroundColor $script:Colors.Warning
            return $false
        }
        
        return $true
    }
    catch {
        Write-Host "ERROR: Failed to initialize LibreHardwareMonitor: $_" -ForegroundColor $script:Colors.Warning
        return $false
    }
}

function Get-CpuPowerConsumption {
    <#
    .SYNOPSIS
        Gets current CPU package power in milliwatts
    #>
    if ($null -eq $script:CpuHardware) {
        return $null
    }
    
    try {
        $script:CpuHardware.Update()
        
        foreach ($sensor in $script:CpuHardware.Sensors) {
            if ($sensor.SensorType -eq [LibreHardwareMonitor.Hardware.SensorType]::Power -and 
                $sensor.Name -eq "CPU Package") {
                if ($sensor.Value) {
                    return [double]$sensor.Value * 1000  # Convert to milliwatts
                }
            }
        }
        
        return $null
    }
    catch {
        return $null
    }
}

function Get-SystemPowerConsumption {
    <#
    .SYNOPSIS
        Gets current system-wide power consumption in milliwatts
    #>
    try {
        $powerCounter = Get-Counter "\Power Meter(_total)\Power" -ErrorAction Stop
        $powerMilliwatts = $powerCounter.CounterSamples[0].CookedValue
        return $powerMilliwatts
    }
    catch {
        return $null
    }
}

function Get-ProcessCpuUtilization {
    <#
    .SYNOPSIS
        Gets CPU utilization for all processes and total CPU
    #>
    try {
        $cpuCounters = Get-Counter "\Process(*)\% Processor Time" -ErrorAction Stop
        
        $processData = @{}
        $totalCpu = 0
        
        foreach ($sample in $cpuCounters.CounterSamples) {
            $processName = $sample.InstanceName
            $cpuPercent = $sample.CookedValue
            
            if ($processName -eq "_total" -or $processName -eq "idle") {
                continue
            }
            
            if ($processData.ContainsKey($processName)) {
                $processData[$processName] += $cpuPercent
            }
            else {
                $processData[$processName] = $cpuPercent
            }
            
            $totalCpu += $cpuPercent
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

function Show-ProcessList {
    <#
    .SYNOPSIS
        Displays a list of running processes with CPU usage for selection
    #>
    param(
        [hashtable]$ProcessCpuData
    )
    
    Clear-Host
    Write-Host "`n=== Per-Process CPU Power Meter ===" -ForegroundColor $script:Colors.Header
    Write-Host "Power Mode: CPU Package Only (RAPL)" -ForegroundColor Green
    Write-Host "CPU: $($script:CpuHardware.Name)" -ForegroundColor $script:Colors.Info
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
    
    Write-Host "`n  0. Refresh list" -ForegroundColor $script:Colors.Info
    Write-Host "  Q. Quit`n" -ForegroundColor $script:Colors.Info
    
    return $processMap
}

function Format-EnergyValue {
    <#
    .SYNOPSIS
        Formats energy value with appropriate unit
    #>
    param([double]$Millijoules)
    
    if ($Millijoules -lt 1000) {
        return "{0:N2} mJ" -f $Millijoules
    }
    elseif ($Millijoules -lt 1000000) {
        return "{0:N2} J" -f ($Millijoules / 1000)
    }
    else {
        return "{0:N6} kJ" -f ($Millijoules / 1000000)
    }
}

function Start-ProcessMonitoring {
    <#
    .SYNOPSIS
        Main monitoring loop for selected process
    #>
    param(
        [string]$ProcessName,
        [int]$IntervalSeconds
    )
    
    $totalEnergyMillijoules = 0
    $totalEnergyMillijoulesSystem = 0
    $measurementCount = 0
    $startTime = Get-Date
    
    Write-Host "`n`n=== Monitoring Process: $ProcessName ===" -ForegroundColor $script:Colors.Header
    Write-Host "Press 'Q' to stop monitoring and return to menu`n" -ForegroundColor $script:Colors.Warning
    
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
        
        # Calculate CPU power allocation
        $processPowerMilliwatts = $cpuPowerMilliwatts * $processUtilizationRatio
        $energyThisPeriodMillijoules = $processPowerMilliwatts * $IntervalSeconds
        $totalEnergyMillijoules += $energyThisPeriodMillijoules
        
        # Calculate system-wide power allocation
        $processPowerMilliwattsSystem = 0
        $energyThisPeriodMillijoulesSystem = 0
        if ($null -ne $systemPowerMilliwatts) {
            $processPowerMilliwattsSystem = $systemPowerMilliwatts * $processUtilizationRatio
            $energyThisPeriodMillijoulesSystem = $processPowerMilliwattsSystem * $IntervalSeconds
            $totalEnergyMillijoulesSystem += $energyThisPeriodMillijoulesSystem
        }
        
        $measurementCount++
        
        # Calculate elapsed time
        $elapsed = (Get-Date) - $startTime
        
        # Display current stats
        Clear-Host
        Write-Host "`n=== Monitoring Process: $ProcessName ===" -ForegroundColor $script:Colors.Header
        Write-Host ("=" * 60) -ForegroundColor $script:Colors.Header
        Write-Host "Mode: Dual Measurement (CPU Package + System Total)" -ForegroundColor Green
        Write-Host "Press 'Q' to stop monitoring and return to menu`n" -ForegroundColor $script:Colors.Warning
        
        Write-Host "Current Measurements:" -ForegroundColor $script:Colors.Info
        Write-Host ("  CPU Package Power (RAPL): {0:N2} W" -f ($cpuPowerMilliwatts / 1000)) -ForegroundColor $script:Colors.Value
        if ($null -ne $systemPowerMilliwatts) {
            Write-Host ("  System Total Power:       {0:N2} W" -f ($systemPowerMilliwatts / 1000)) -ForegroundColor $script:Colors.Value
        }
        Write-Host ("  Process CPU Usage:        {0:N2}%" -f $processCpu) -ForegroundColor $script:Colors.Value
        Write-Host ("  Total CPU Usage:          {0:N2}%" -f $totalCpu) -ForegroundColor $script:Colors.Value
        Write-Host ("  Process CPU Power:        {0:N2} W" -f ($processPowerMilliwatts / 1000)) -ForegroundColor $script:Colors.Value
        if ($null -ne $systemPowerMilliwatts) {
            Write-Host ("  Process System Power:     {0:N2} W" -f ($processPowerMilliwattsSystem / 1000)) -ForegroundColor $script:Colors.Value
        }
        
        Write-Host "`nAccumulated Statistics (CPU Only - RAPL):" -ForegroundColor $script:Colors.Info
        Write-Host ("  Total Energy Consumed:    {0}" -f (Format-EnergyValue $totalEnergyMillijoules)) -ForegroundColor $script:Colors.Value
        Write-Host ("  Average Power:            {0:N2} W" -f (($totalEnergyMillijoules / 1000) / $elapsed.TotalSeconds)) -ForegroundColor $script:Colors.Value
        
        if ($null -ne $systemPowerMilliwatts) {
            Write-Host "`nAccumulated Statistics (System Total):" -ForegroundColor $script:Colors.Info
            Write-Host ("  Total Energy Consumed:    {0}" -f (Format-EnergyValue $totalEnergyMillijoulesSystem)) -ForegroundColor $script:Colors.Value
            Write-Host ("  Average Power:            {0:N2} W" -f (($totalEnergyMillijoulesSystem / 1000) / $elapsed.TotalSeconds)) -ForegroundColor $script:Colors.Value
        }
        
        Write-Host "`nGeneral Statistics:" -ForegroundColor $script:Colors.Info
        Write-Host ("  Monitoring Duration:      {0:hh\:mm\:ss}" -f $elapsed) -ForegroundColor $script:Colors.Value
        Write-Host ("  Measurements Taken:       {0}" -f $measurementCount) -ForegroundColor $script:Colors.Value
        
        Start-Sleep -Seconds $IntervalSeconds
    }
    
    # Final summary
    Write-Host ("Process:              {0}" -f $ProcessName) -ForegroundColor $script:Colors.ProcessName
    Write-Host "`nCPU Only (RAPL):" -ForegroundColor $script:Colors.Info
    Write-Host ("  Total Energy:       {0}" -f (Format-EnergyValue $totalEnergyMillijoules)) -ForegroundColor $script:Colors.Value
    Write-Host ("  Average Power:      {0:N2} W" -f (($totalEnergyMillijoules / 1000) / $elapsed.TotalSeconds)) -ForegroundColor $script:Colors.Value
    
    if ($null -ne $systemPowerMilliwatts -and $totalEnergyMillijoulesSystem -gt 0) {
        Write-Host "`nSystem Total:" -ForegroundColor $script:Colors.Info
        Write-Host ("  Total Energy:       {0}" -f (Format-EnergyValue $totalEnergyMillijoulesSystem)) -ForegroundColor $script:Colors.Value
        Write-Host ("  Average Power:      {0:N2} W" -f (($totalEnergyMillijoulesSystem / 1000) / $elapsed.TotalSeconds)) -ForegroundColor $script:Colors.Value
    }
    
    Write-Host "`nMonitoring Duration:  {0:hh\:mm\:ss}" -f $elapsed -ForegroundColor $script:Colors.Value
    Write-Host "`nNote: CPU measurements include CPU package only (cores + cache)" -ForegroundColor $script:Colors.Info
    Write-Host "      System measurements include all components (CPU + GPU + display + etc)" -ForegroundColor $script:Colors.Info
    Write-Host "`nPress any key to return to menu..." -ForegroundColor $script:Colors.Info
    $null = [Console]::ReadKey($true)
}

function Start-PowerMeterApp {
    param([int]$IntervalSeconds)
    
    Write-Host "`nInitializing Per-Process CPU Power Meter..." -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    
    $initialized = Initialize-LibreHardwareMonitor
    
    if (-not $initialized) {
        Write-Host "`nFailed to initialize. Exiting..." -ForegroundColor $script:Colors.Warning
        return
    }
    
    Write-Host "[OK] CPU power monitoring ready" -ForegroundColor Green
    Write-Host "`nPress any key to continue..." -ForegroundColor $script:Colors.Info
    $null = [Console]::ReadKey($true)
    
    while ($true) {
        $cpuData = Get-ProcessCpuUtilization
        
        if ($null -eq $cpuData) {
            Write-Host "Error reading CPU data. Exiting..." -ForegroundColor $script:Colors.Warning
            return
        }
        
        $processMap = Show-ProcessList -ProcessCpuData $cpuData.ProcessData
        
        Write-Host "Enter selection: " -NoNewline -ForegroundColor $script:Colors.Info
        $selection = Read-Host
        
        if ($selection -eq 'Q' -or $selection -eq 'q') {
            Write-Host "`nExiting..." -ForegroundColor $script:Colors.Info
            break
        }
        
        if ($selection -eq '0') {
            continue
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
    
    # Cleanup
    if ($null -ne $script:Computer) {
        $script:Computer.Close()
    }
}

# Check admin privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "`nERROR: This script requires Administrator privileges" -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again.`n" -ForegroundColor Yellow
    exit 1
}

# Start the application
Start-PowerMeterApp -IntervalSeconds $MeasurementIntervalSeconds
