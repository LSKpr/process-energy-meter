<#
.SYNOPSIS
    Per-Process CPU Power Meter
.DESCRIPTION
    Monitors and calculates power consumption for selected processes based on 
    their CPU utilization and system power draw.
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

function Get-SystemPowerConsumption {
    <#
    .SYNOPSIS
        Gets current system power consumption in milliwatts
    #>
    try {
        $powerCounter = Get-Counter "\Power Meter(_total)\Power" -ErrorAction Stop
        $powerMilliwatts = $powerCounter.CounterSamples[0].CookedValue
        return $powerMilliwatts
    }
    catch {
        Write-Warning "Unable to read Power Meter counter. Make sure your system supports it."
        return $null
    }
}

function Get-ProcessCpuUtilization {
    <#
    .SYNOPSIS
        Gets CPU utilization for all processes and total CPU
    #>
    try {
        # Get all processor time counters
        $cpuCounters = Get-Counter "\Process(*)\% Processor Time" -ErrorAction Stop
        
        $processData = @{}
        $totalCpu = 0
        
        foreach ($sample in $cpuCounters.CounterSamples) {
            $processName = $sample.InstanceName
            $cpuPercent = $sample.CookedValue
            
            # Skip _total and idle
            if ($processName -eq "_total" -or $processName -eq "idle") {
                continue
            }
            
            # Aggregate processes with same name
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
    Write-Host "==================================" -ForegroundColor $script:Colors.Header
    Write-Host "`nSelect a process to monitor:`n" -ForegroundColor $script:Colors.Info
    
    # Sort by CPU usage and get top processes
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
    $measurementCount = 0
    $startTime = Get-Date
    
    Write-Host "`n`n=== Monitoring Process: $ProcessName ===" -ForegroundColor $script:Colors.Header
    Write-Host "Press 'Q' to stop monitoring and return to menu`n" -ForegroundColor $script:Colors.Warning
    
    while ($true) {
        # Check for quit key
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq 'Q') {
                break
            }
        }
        
        # Get current measurements
        $cpuData = Get-ProcessCpuUtilization
        $powerMilliwatts = Get-SystemPowerConsumption
        
        if ($null -eq $cpuData -or $null -eq $powerMilliwatts) {
            Write-Host "Error collecting data. Retrying..." -ForegroundColor $script:Colors.Warning
            Start-Sleep -Seconds 1
            continue
        }
        
        # Calculate process power
        $processCpu = 0
        if ($cpuData.ProcessData.ContainsKey($ProcessName)) {
            $processCpu = $cpuData.ProcessData[$ProcessName]
        }
        
        $totalCpu = $cpuData.TotalCpu
        
        # Avoid division by zero
        if ($totalCpu -gt 0) {
            $processUtilizationRatio = $processCpu / $totalCpu
        }
        else {
            $processUtilizationRatio = 0
        }
        
        # Energy = Power × Time
        # Power for process = Total Power × (Process CPU / Total CPU)
        $processPowerMilliwatts = $powerMilliwatts * $processUtilizationRatio
        $energyThisPeriodMillijoules = $processPowerMilliwatts * $IntervalSeconds
        $totalEnergyMillijoules += $energyThisPeriodMillijoules
        $measurementCount++
        
        # Calculate elapsed time
        $elapsed = (Get-Date) - $startTime
        
        # Display current stats
        Clear-Host
        Write-Host "`n=== Monitoring Process: $ProcessName ===" -ForegroundColor $script:Colors.Header
        Write-Host ("=" * 60) -ForegroundColor $script:Colors.Header
        Write-Host "Press 'Q' to stop monitoring and return to menu`n" -ForegroundColor $script:Colors.Warning
        
        Write-Host "Current Measurements:" -ForegroundColor $script:Colors.Info
        Write-Host ("  System Power:          {0:N2} W" -f ($powerMilliwatts / 1000)) -ForegroundColor $script:Colors.Value
        Write-Host ("  Process CPU Usage:     {0:N2}%" -f $processCpu) -ForegroundColor $script:Colors.Value
        Write-Host ("  Total CPU Usage:       {0:N2}%" -f $totalCpu) -ForegroundColor $script:Colors.Value
        Write-Host ("  Process Power Share:   {0:N2} W" -f ($processPowerMilliwatts / 1000)) -ForegroundColor $script:Colors.Value
        
        Write-Host "`nAccumulated Statistics:" -ForegroundColor $script:Colors.Info
        Write-Host ("  Total Energy Consumed: {0}" -f (Format-EnergyValue $totalEnergyMillijoules)) -ForegroundColor $script:Colors.Value
        Write-Host ("  Average Power:         {0:N2} W" -f (($totalEnergyMillijoules / 1000) / $elapsed.TotalSeconds)) -ForegroundColor $script:Colors.Value
        Write-Host ("  Monitoring Duration:   {0:hh\:mm\:ss}" -f $elapsed) -ForegroundColor $script:Colors.Value
        Write-Host ("  Measurements Taken:    {0}" -f $measurementCount) -ForegroundColor $script:Colors.Value
        
        Start-Sleep -Seconds $IntervalSeconds
    }
    
    # Final summary
    Write-Host "`n`n=== Monitoring Summary ===" -ForegroundColor $script:Colors.Header
    Write-Host ("Process:              {0}" -f $ProcessName) -ForegroundColor $script:Colors.ProcessName
    Write-Host ("Total Energy:         {0}" -f (Format-EnergyValue $totalEnergyMillijoules)) -ForegroundColor $script:Colors.Value
    Write-Host ("Monitoring Duration:  {0:hh\:mm\:ss}" -f $elapsed) -ForegroundColor $script:Colors.Value
    Write-Host ("Average Power:        {0:N2} W" -f (($totalEnergyMillijoules / 1000) / $elapsed.TotalSeconds)) -ForegroundColor $script:Colors.Value
    Write-Host "`nPress any key to return to menu..." -ForegroundColor $script:Colors.Info
    $null = [Console]::ReadKey($true)
}

# Main program loop
function Start-PowerMeterApp {
    param([int]$IntervalSeconds)
    
    # Check if Power Meter is available
    $testPower = Get-SystemPowerConsumption
    if ($null -eq $testPower) {
        Write-Host "`nERROR: Power Meter counters are not available on this system." -ForegroundColor $script:Colors.Warning
        Write-Host "This tool requires a system with Power Meter support (typically modern laptops/devices)." -ForegroundColor $script:Colors.Warning
        return
    }
    
    while ($true) {
        # Get current CPU data
        $cpuData = Get-ProcessCpuUtilization
        
        if ($null -eq $cpuData) {
            Write-Host "Error reading CPU data. Exiting..." -ForegroundColor $script:Colors.Warning
            return
        }
        
        # Show process selection menu
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
}

# Start the application
Start-PowerMeterApp -IntervalSeconds $MeasurementIntervalSeconds
