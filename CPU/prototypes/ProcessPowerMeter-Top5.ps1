<#
.SYNOPSIS
    Top 5 Process Power Monitor
.DESCRIPTION
    Continuously monitors ALL processes and displays the top 5 by accumulated
    power consumption using LibreHardwareMonitor's RAPL interface.
    
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
    Highlight = 'Magenta'
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
        # Treat 0 as unavailable
        if ($powerMilliwatts -le 0) {
            return $null
        }
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

function Format-EnergyValue {
    param([double]$Millijoules)
    
    if ($Millijoules -gt 1000000) {
        return "{0:N2} kJ" -f ($Millijoules / 1000000)
    }
    elseif ($Millijoules -gt 1000) {
        return "{0:N2} J" -f ($Millijoules / 1000)
    }
    else {
        return "{0:N0} mJ" -f $Millijoules
    }
}

function Start-Top5Monitor {
    param([int]$IntervalSeconds)
    
    # Hash table to accumulate energy per process
    $processEnergyAccumulator = @{}
    
    $measurementCount = 0
    $startTime = Get-Date
    
    Write-Host "`n=== Top 5 Process Power Monitor ===" -ForegroundColor $script:Colors.Header
    Write-Host "CPU: $($script:CpuHardware.Name)" -ForegroundColor $script:Colors.Info
    Write-Host "Monitoring all processes... Press 'Q' to stop`n" -ForegroundColor $script:Colors.Warning
    Start-Sleep -Seconds 1
    
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
        $cpuPowerMilliwatts = Get-CpuPowerConsumption
        $systemPowerMilliwatts = Get-SystemPowerConsumption
        
        if ($null -eq $cpuData -or $null -eq $cpuPowerMilliwatts) {
            Write-Host "Error collecting data. Retrying..." -ForegroundColor $script:Colors.Warning
            Start-Sleep -Seconds 1
            continue
        }
        
        $totalCpu = $cpuData.TotalCpu
        
        # Calculate power for each process
        if ($totalCpu -gt 0) {
            foreach ($process in $cpuData.ProcessData.GetEnumerator()) {
                $processName = $process.Key
                $processCpu = $process.Value
                
                $cpuRatio = $processCpu / $totalCpu
                
                # CPU power allocation
                $processPowerCpu = $cpuPowerMilliwatts * $cpuRatio
                $energyCpu = $processPowerCpu * $IntervalSeconds
                
                # System power allocation (if available)
                $energySystem = 0
                if ($null -ne $systemPowerMilliwatts) {
                    $processPowerSystem = $systemPowerMilliwatts * $cpuRatio
                    $energySystem = $processPowerSystem * $IntervalSeconds
                }
                
                # Accumulate energy
                if (-not $processEnergyAccumulator.ContainsKey($processName)) {
                    $processEnergyAccumulator[$processName] = @{
                        CpuEnergy = 0
                        SystemEnergy = 0
                        LastSeenCpu = 0
                    }
                }
                
                $processEnergyAccumulator[$processName].CpuEnergy += $energyCpu
                $processEnergyAccumulator[$processName].SystemEnergy += $energySystem
                $processEnergyAccumulator[$processName].LastSeenCpu = $processCpu
            }
        }
        
        $measurementCount++
        $elapsed = (Get-Date) - $startTime
        
        # Get top 5 processes by CPU energy
        $top5 = $processEnergyAccumulator.GetEnumerator() | 
            Sort-Object { $_.Value.CpuEnergy } -Descending |
            Select-Object -First 5
        
        # Display
        Clear-Host
        Write-Host "`n=== Top 5 Process Power Monitor ===" -ForegroundColor $script:Colors.Header
        Write-Host ("=" * 70) -ForegroundColor $script:Colors.Header
        Write-Host "CPU: $($script:CpuHardware.Name)" -ForegroundColor $script:Colors.Info
        Write-Host "Press 'Q' to stop monitoring`n" -ForegroundColor $script:Colors.Warning
        
        Write-Host "Current System Power:" -ForegroundColor $script:Colors.Info
        Write-Host ("  CPU Package (RAPL):  {0:N2} W" -f ($cpuPowerMilliwatts / 1000)) -ForegroundColor $script:Colors.Value
        if ($null -ne $systemPowerMilliwatts) {
            Write-Host ("  System Total:        {0:N2} W" -f ($systemPowerMilliwatts / 1000)) -ForegroundColor $script:Colors.Value
        }
        else {
            Write-Host ("  System Total:        Not Available") -ForegroundColor Yellow
        }
        
        Write-Host "`nTop 5 Processes by Accumulated Energy:" -ForegroundColor $script:Colors.Highlight
        Write-Host ("{0,-30} {1,12} {2,12} {3,10}" -f "Process", "CPU Energy", "System Energy", "Current CPU") -ForegroundColor $script:Colors.Header
        Write-Host ("-" * 70) -ForegroundColor $script:Colors.Header
        
        $rank = 1
        foreach ($proc in $top5) {
            $cpuEnergy = Format-EnergyValue $proc.Value.CpuEnergy
            $systemEnergy = if ($proc.Value.SystemEnergy -gt 0) { Format-EnergyValue $proc.Value.SystemEnergy } else { "N/A" }
            $currentCpu = "{0:N2}%" -f $proc.Value.LastSeenCpu
            
            $color = switch ($rank) {
                1 { 'Yellow' }
                2 { 'Green' }
                3 { 'Cyan' }
                default { 'White' }
            }
            
            Write-Host ("{0}. {1,-28} {2,12} {3,12} {4,10}" -f $rank, $proc.Key, $cpuEnergy, $systemEnergy, $currentCpu) -ForegroundColor $color
            $rank++
        }
        
        Write-Host "`nSession Statistics:" -ForegroundColor $script:Colors.Info
        Write-Host ("  Duration:              {0:hh\:mm\:ss}" -f $elapsed) -ForegroundColor $script:Colors.Value
        Write-Host ("  Measurements:          {0}" -f $measurementCount) -ForegroundColor $script:Colors.Value
        Write-Host ("  Tracked Processes:     {0}" -f $processEnergyAccumulator.Count) -ForegroundColor $script:Colors.Value
        
        # Calculate total energy consumed by all processes
        $totalCpuEnergy = 0
        $totalSystemEnergy = 0
        foreach ($value in $processEnergyAccumulator.Values) {
            $totalCpuEnergy += $value.CpuEnergy
            $totalSystemEnergy += $value.SystemEnergy
        }
        Write-Host ("  Total CPU Energy:      {0}" -f (Format-EnergyValue $totalCpuEnergy)) -ForegroundColor $script:Colors.Value
        
        if ($null -ne $systemPowerMilliwatts) {
            Write-Host ("  Total System Energy:   {0}" -f (Format-EnergyValue $totalSystemEnergy)) -ForegroundColor $script:Colors.Value
        }
        
        Start-Sleep -Seconds $IntervalSeconds
    }
    
    # Final summary
    Write-Host "`n`n=== Final Summary ===" -ForegroundColor $script:Colors.Header
    Write-Host ("Duration:              {0:hh\:mm\:ss}" -f $elapsed) -ForegroundColor $script:Colors.Value
    Write-Host ("Total Measurements:    {0}" -f $measurementCount) -ForegroundColor $script:Colors.Value
    Write-Host ("Tracked Processes:     {0}" -f $processEnergyAccumulator.Count) -ForegroundColor $script:Colors.Value
    
    Write-Host "`nTop 10 Processes by Total CPU Energy:" -ForegroundColor $script:Colors.Highlight
    Write-Host ("{0,-35} {1,15} {2,15}" -f "Process", "CPU Energy", "System Energy") -ForegroundColor $script:Colors.Header
    Write-Host ("-" * 70) -ForegroundColor $script:Colors.Header
    
    $top10Final = $processEnergyAccumulator.GetEnumerator() | 
        Sort-Object { $_.Value.CpuEnergy } -Descending |
        Select-Object -First 10
    
    $rank = 1
    foreach ($proc in $top10Final) {
        $cpuEnergy = Format-EnergyValue $proc.Value.CpuEnergy
        $systemEnergy = if ($proc.Value.SystemEnergy -gt 0) { Format-EnergyValue $proc.Value.SystemEnergy } else { "N/A" }
        
        Write-Host ("{0,2}. {1,-32} {2,15} {3,15}" -f $rank, $proc.Key, $cpuEnergy, $systemEnergy) -ForegroundColor $script:Colors.ProcessName
        $rank++
    }
    
    Write-Host "`nPress any key to exit..." -ForegroundColor $script:Colors.Info
    $null = [Console]::ReadKey($true)
}

# Main execution
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "`nERROR: This script requires Administrator privileges!" -ForegroundColor $script:Colors.Warning
    Write-Host "Please run PowerShell as Administrator and try again.`n" -ForegroundColor $script:Colors.Info
    exit 1
}

Write-Host "`nInitializing Top 5 Process Power Monitor..." -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor Cyan

$initialized = Initialize-LibreHardwareMonitor

if (-not $initialized) {
    Write-Host "`nFailed to initialize. Exiting..." -ForegroundColor $script:Colors.Warning
    exit 1
}

Write-Host "[OK] CPU power monitoring ready (RAPL)" -ForegroundColor Green

# Check system power availability
$testSystemPower = Get-SystemPowerConsumption
if ($null -eq $testSystemPower) {
    Write-Host "[!] System total power unavailable (Power Meter not supported)" -ForegroundColor Yellow
    Write-Host "    Will show CPU power only" -ForegroundColor Yellow
}
else {
    Write-Host "[OK] System total power available" -ForegroundColor Green
}

Write-Host "`nPress any key to start monitoring..." -ForegroundColor $script:Colors.Info
$null = [Console]::ReadKey($true)

Start-Top5Monitor -IntervalSeconds $MeasurementIntervalSeconds

# Cleanup
if ($null -ne $script:Computer) {
    $script:Computer.Close()
}
