<#
.SYNOPSIS
    Per-Process Power Meter with Memory-Weighted Allocation
.DESCRIPTION
    Monitors power consumption with intelligent allocation:
    - CPU power allocated by CPU usage
    - Memory power allocated by memory usage (working set)
    - Remaining system power allocated proportionally
    
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

function Get-MemoryPowerConsumption {
    <#
    .SYNOPSIS
        Gets DRAM power from RAPL if available
    #>
    if ($null -eq $script:CpuHardware) {
        return $null
    }
    
    try {
        $script:CpuHardware.Update()
        
        # Look for DRAM/Memory power sensor
        $memorySensor = $script:CpuHardware.Sensors | Where-Object {
            $_.SensorType -eq [LibreHardwareMonitor.Hardware.SensorType]::Power -and 
            ($_.Name -like "*DRAM*" -or $_.Name -like "*Memory*")
        } | Select-Object -First 1
        
        if ($null -ne $memorySensor -and $null -ne $memorySensor.Value) {
            return [double]$memorySensor.Value * 1000
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
        # Treat 0 as unavailable (common on laptops without proper power meter support)
        if ($powerMilliwatts -le 0) {
            return $null
        }
        return $powerMilliwatts
    }
    catch {
        return $null
    }
}

function Get-ProcessResourceUtilization {
    <#
    .SYNOPSIS
        Gets CPU and memory utilization for all processes
    #>
    try {
        # Get CPU usage
        $cpuCounters = Get-Counter -Counter "\Process(*)\% Processor Time" -ErrorAction Stop
        
        $processData = @{}
        $totalCpu = 0
        $totalMemoryMB = 0
        
        foreach ($sample in $cpuCounters.CounterSamples) {
            $processName = $sample.InstanceName
            $cpuValue = $sample.CookedValue
            
            if ($processName -eq '_total' -or $processName -eq 'idle') {
                continue
            }
            
            if (-not $processData.ContainsKey($processName)) {
                $processData[$processName] = @{
                    CPU = 0
                    MemoryMB = 0
                }
            }
            
            $processData[$processName].CPU += $cpuValue
            $totalCpu += $cpuValue
        }
        
        # Get memory usage for all processes
        $processes = Get-Process -ErrorAction SilentlyContinue
        
        foreach ($proc in $processes) {
            $processName = $proc.Name
            
            if ($processData.ContainsKey($processName)) {
                $memoryMB = $proc.WorkingSet64 / 1MB
                $processData[$processName].MemoryMB += $memoryMB
                $totalMemoryMB += $memoryMB
            }
        }
        
        return @{
            ProcessData = $processData
            TotalCpu = $totalCpu
            TotalMemoryMB = $totalMemoryMB
        }
    }
    catch {
        Write-Warning "Error reading process data: $_"
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

function Show-ProcessList {
    param([hashtable]$ProcessData)
    
    Clear-Host
    Write-Host "`n=== Per-Process Power Meter (Memory-Weighted) ===" -ForegroundColor $script:Colors.Header
    Write-Host "Power Mode: CPU + Memory + System Breakdown" -ForegroundColor Green
    Write-Host "CPU: $($script:CpuHardware.Name)" -ForegroundColor $script:Colors.Info
    Write-Host "================================================" -ForegroundColor $script:Colors.Header
    Write-Host "`nSelect a process to monitor:`n" -ForegroundColor $script:Colors.Info
    
    $sortedProcesses = $ProcessData.GetEnumerator() | 
        Where-Object { $_.Value.CPU -gt 0 -or $_.Value.MemoryMB -gt 50 } |
        Sort-Object -Property { $_.Value.CPU } -Descending |
        Select-Object -First 30
    
    $index = 1
    $processMap = @{}
    
    foreach ($proc in $sortedProcesses) {
        $processMap[$index] = $proc.Key
        $cpuDisplay = "{0:N2}%" -f $proc.Value.CPU
        $memDisplay = "{0:N0} MB" -f $proc.Value.MemoryMB
        Write-Host ("{0,3}. {1,-25} CPU: {2,-8} RAM: {3}" -f $index, $proc.Key, $cpuDisplay, $memDisplay) -ForegroundColor $script:Colors.ProcessName
        $index++
    }
    
    Write-Host "`n  0. Return to main menu" -ForegroundColor $script:Colors.Warning
    Write-Host "  Q. Quit`n" -ForegroundColor $script:Colors.Warning
    
    return $processMap
}

function Start-ProcessMonitoring {
    param(
        [string]$ProcessName,
        [int]$IntervalSeconds
    )
    
    # Energy accumulators
    $totalEnergyCpu = 0
    $totalEnergyMemory = 0
    $totalEnergyOther = 0
    $totalEnergySystem = 0
    
    $measurementCount = 0
    $startTime = Get-Date
    
    Write-Host "`nStarting monitoring for: $ProcessName" -ForegroundColor $script:Colors.Info
    Write-Host "Press 'Q' to stop...`n" -ForegroundColor $script:Colors.Warning
    Start-Sleep -Seconds 1
    
    # Estimate memory power as % of total (rough estimate: 10-25% of system power)
    $memoryPowerRatio = 0.15  # Conservative 15% estimate
    
    while ($true) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq 'Q') {
                break
            }
        }
        
        $resourceData = Get-ProcessResourceUtilization
        $cpuPowerMilliwatts = Get-CpuPowerConsumption
        $memoryPowerMilliwatts = Get-MemoryPowerConsumption
        $systemPowerMilliwatts = Get-SystemPowerConsumption
        
        if ($null -eq $resourceData -or $null -eq $cpuPowerMilliwatts) {
            Write-Host "Error collecting data. Retrying..." -ForegroundColor $script:Colors.Warning
            Start-Sleep -Seconds 1
            continue
        }
        
        # Get process stats
        $processCpu = 0
        $processMemoryMB = 0
        
        if ($resourceData.ProcessData.ContainsKey($ProcessName)) {
            $processCpu = $resourceData.ProcessData[$ProcessName].CPU
            $processMemoryMB = $resourceData.ProcessData[$ProcessName].MemoryMB
        }
        
        $totalCpu = $resourceData.TotalCpu
        $totalMemoryMB = $resourceData.TotalMemoryMB
        
        # Calculate ratios
        $cpuRatio = if ($totalCpu -gt 0) { $processCpu / $totalCpu } else { 0 }
        $memoryRatio = if ($totalMemoryMB -gt 0) { $processMemoryMB / $totalMemoryMB } else { 0 }
        
        # Allocate CPU power by CPU usage
        $processPowerCpu = $cpuPowerMilliwatts * $cpuRatio
        $energyCpu = $processPowerCpu * $IntervalSeconds
        $totalEnergyCpu += $energyCpu
        
        # Allocate memory power
        $estimatedMemoryPower = 0
        $processPowerMemory = 0
        $energyMemory = 0
        $memoryPowerSource = "Estimated"
        
        if ($null -ne $memoryPowerMilliwatts) {
            # Use actual DRAM RAPL measurement
            $estimatedMemoryPower = $memoryPowerMilliwatts
            $memoryPowerSource = "RAPL"
        }
        elseif ($null -ne $systemPowerMilliwatts) {
            # Estimate as ~15% of system power
            $estimatedMemoryPower = $systemPowerMilliwatts * $memoryPowerRatio
            $memoryPowerSource = "Estimated"
        }
        
        if ($estimatedMemoryPower -gt 0) {
            $processPowerMemory = $estimatedMemoryPower * $memoryRatio
            $energyMemory = $processPowerMemory * $IntervalSeconds
            $totalEnergyMemory += $energyMemory
        }
        
        # Calculate remaining power (if system power available)
        $processPowerOther = 0
        $energyOther = 0
        $totalProcessPower = $processPowerCpu + $processPowerMemory
        
        if ($null -ne $systemPowerMilliwatts) {
            # Remaining power (GPU, Display, Storage, etc.) - allocate by CPU ratio as proxy
            $remainingPower = $systemPowerMilliwatts - $cpuPowerMilliwatts - $estimatedMemoryPower
            if ($remainingPower -lt 0) { $remainingPower = 0 }
            
            $processPowerOther = $remainingPower * $cpuRatio
            $energyOther = $processPowerOther * $IntervalSeconds
            $totalEnergyOther += $energyOther
            
            # Total system allocation
            $totalProcessPower = $processPowerCpu + $processPowerMemory + $processPowerOther
            $energySystem = $totalProcessPower * $IntervalSeconds
            $totalEnergySystem += $energySystem
        }
        
        $measurementCount++
        $elapsed = (Get-Date) - $startTime
        
        # Display
        Clear-Host
        Write-Host "`n=== Monitoring Process: $ProcessName ===" -ForegroundColor $script:Colors.Header
        Write-Host ("=" * 70) -ForegroundColor $script:Colors.Header
        Write-Host "Mode: Component-Based Power Allocation" -ForegroundColor Green
        Write-Host "Press 'Q' to stop monitoring`n" -ForegroundColor $script:Colors.Warning
        
        Write-Host "Current Power Breakdown:" -ForegroundColor $script:Colors.Info
        Write-Host ("  CPU Package Power (RAPL):     {0:N2} W" -f ($cpuPowerMilliwatts / 1000)) -ForegroundColor $script:Colors.Value
        if ($estimatedMemoryPower -gt 0) {
            Write-Host ("  Memory Power ({0}):        {1:N2} W" -f $memoryPowerSource, ($estimatedMemoryPower / 1000)) -ForegroundColor $script:Colors.Value
        }
        if ($null -ne $systemPowerMilliwatts) {
            Write-Host ("  System Total Power:            {0:N2} W" -f ($systemPowerMilliwatts / 1000)) -ForegroundColor $script:Colors.Value
        }
        else {
            Write-Host ("  System Total Power:            Not Available" -f $memoryPowerSource) -ForegroundColor Yellow
        }
        
        Write-Host "`nProcess Utilization:" -ForegroundColor $script:Colors.Info
        Write-Host ("  CPU Usage:                     {0:N2}% of {1:N2}%" -f $processCpu, $totalCpu) -ForegroundColor $script:Colors.Value
        Write-Host ("  Memory Usage:                  {0:N0} MB of {1:N0} MB ({2:P1})" -f $processMemoryMB, $totalMemoryMB, $memoryRatio) -ForegroundColor $script:Colors.Value
        
        Write-Host "`nAllocated Power to Process:" -ForegroundColor $script:Colors.Info
        Write-Host ("  CPU Power:                     {0:N3} W" -f ($processPowerCpu / 1000)) -ForegroundColor $script:Colors.Value
        if ($null -ne $systemPowerMilliwatts) {
            Write-Host ("  Memory Power:                  {0:N3} W" -f ($processPowerMemory / 1000)) -ForegroundColor $script:Colors.Value
            Write-Host ("  Other (GPU/Display/Storage):   {0:N3} W" -f ($processPowerOther / 1000)) -ForegroundColor $script:Colors.Value
            Write-Host ("  Total Process Power:           {0:N3} W" -f ($totalProcessPower / 1000)) -ForegroundColor Yellow
        }
        
        Write-Host "`nAccumulated Energy:" -ForegroundColor $script:Colors.Info
        Write-Host ("  CPU Energy:                    {0}" -f (Format-EnergyValue $totalEnergyCpu)) -ForegroundColor $script:Colors.Value
        
        if ($totalEnergySystem -gt 0) {
            Write-Host ("  Memory Energy:                 {0}" -f (Format-EnergyValue $totalEnergyMemory)) -ForegroundColor $script:Colors.Value
            Write-Host ("  Other Energy:                  {0}" -f (Format-EnergyValue $totalEnergyOther)) -ForegroundColor $script:Colors.Value
            Write-Host ("  Total Energy:                  {0}" -f (Format-EnergyValue $totalEnergySystem)) -ForegroundColor Yellow
        }
        
        Write-Host "`nSession Statistics:" -ForegroundColor $script:Colors.Info
        Write-Host ("  Duration:                      {0:hh\:mm\:ss}" -f $elapsed) -ForegroundColor $script:Colors.Value
        Write-Host ("  Measurements:                  {0}" -f $measurementCount) -ForegroundColor $script:Colors.Value
        if ($totalEnergySystem -gt 0) {
            Write-Host ("  Average Total Power:           {0:N3} W" -f (($totalEnergySystem / 1000) / $elapsed.TotalSeconds)) -ForegroundColor Yellow
        }
        else {
            Write-Host ("  Average CPU Power:             {0:N3} W" -f (($totalEnergyCpu / 1000) / $elapsed.TotalSeconds)) -ForegroundColor Yellow
        }
        
        Start-Sleep -Seconds $IntervalSeconds
    }
    
    # Final summary
    Write-Host "`n`n=== Monitoring Summary ===" -ForegroundColor $script:Colors.Header
    Write-Host ("Process:              {0}" -f $ProcessName) -ForegroundColor $script:Colors.ProcessName
    
    if ($totalEnergySystem -gt 0) {
        Write-Host "`nComponent Breakdown:" -ForegroundColor $script:Colors.Info
        Write-Host ("  CPU Energy:         {0}" -f (Format-EnergyValue $totalEnergyCpu)) -ForegroundColor $script:Colors.Value
        Write-Host ("  Memory Energy:      {0}" -f (Format-EnergyValue $totalEnergyMemory)) -ForegroundColor $script:Colors.Value
        Write-Host ("  Other Energy:       {0}" -f (Format-EnergyValue $totalEnergyOther)) -ForegroundColor $script:Colors.Value
        Write-Host ("  Total Energy:       {0}" -f (Format-EnergyValue $totalEnergySystem)) -ForegroundColor Yellow
        
        $avgPowerW = ($totalEnergySystem / 1000) / $elapsed.TotalSeconds
        Write-Host "`nAverage Power:        {0:N3} W" -f $avgPowerW -ForegroundColor Yellow
    }
    else {
        Write-Host "`nCPU Energy:           {0}" -f (Format-EnergyValue $totalEnergyCpu) -ForegroundColor Yellow
        
        $avgPowerW = ($totalEnergyCpu / 1000) / $elapsed.TotalSeconds
        Write-Host "Average CPU Power:    {0:N3} W" -f $avgPowerW -ForegroundColor Yellow
        Write-Host "`nNote: System total power unavailable on this system" -ForegroundColor Yellow
        Write-Host "      Showing CPU package power only" -ForegroundColor Yellow
    }
    
    Write-Host ("Duration:             {0:hh\:mm\:ss}" -f $elapsed) -ForegroundColor $script:Colors.Value
    Write-Host "`nPress any key to return to menu..." -ForegroundColor $script:Colors.Info
    $null = [Console]::ReadKey($true)
}

function Start-PowerMeterApp {
    param([int]$IntervalSeconds)
    
    Write-Host "`nInitializing Per-Process Power Meter..." -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    
    $initialized = Initialize-LibreHardwareMonitor
    
    if (-not $initialized) {
        Write-Host "`nFailed to initialize. Exiting..." -ForegroundColor $script:Colors.Warning
        return
    }
    
    Write-Host "[OK] CPU power monitoring ready (RAPL)" -ForegroundColor Green
    
    # Check system power availability
    $testSystemPower = Get-SystemPowerConsumption
    if ($null -eq $testSystemPower) {
        Write-Host "[!] System total power unavailable (Power Meter not supported)" -ForegroundColor Yellow
        Write-Host "    Will measure CPU power only" -ForegroundColor Yellow
    }
    else {
        Write-Host "[OK] System total power available" -ForegroundColor Green
    }
    
    Write-Host "`nPress any key to continue..." -ForegroundColor $script:Colors.Info
    $null = [Console]::ReadKey($true)
    
    while ($true) {
        $resourceData = Get-ProcessResourceUtilization
        
        if ($null -eq $resourceData) {
            Write-Host "Error reading process data. Exiting..." -ForegroundColor $script:Colors.Warning
            return
        }
        
        $processMap = Show-ProcessList -ProcessData $resourceData.ProcessData
        
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
