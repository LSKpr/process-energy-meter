#----------------------------------------------
#region Initialization
#----------------------------------------------

$script:Computer = $null
$script:CpuHardware = $null

$script:ProcessEnergyData = @{}
$script:MeasurementHistory = New-Object System.Collections.ArrayList
$script:MaxHistorySize = 100
$script:MeasurementCount = 0
$script:CurrentCpuPower = 0
$script:CurrentSystemPower = 0
$script:CurrentTotalCpuPercent = 0

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
            #find cpu
            if ($hardware.HardwareType -eq [LibreHardwareMonitor.Hardware.HardwareType]::Cpu) {
                $script:CpuHardware = $hardware
                return $true
            }
        }
        
        return $false
    }
    catch {
        Write-Host "ERROR: Failed to initialize LibreHardwareMonitor: $_" -ForegroundColor $script:Colors.Warning
        return $false
    }
}

#endregion

#----------------------------------------------
#region Core Sampling
#----------------------------------------------

function Get-CpuPowerConsumption {
    if ($null -eq $script:CpuHardware) { return $null }
    
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
    try {
        $powerCounter = Get-Counter "\Power Meter(_total)\Power" -ErrorAction Stop
        $powerMilliwatts = $powerCounter.CounterSamples[0].CookedValue
        if ($powerMilliwatts -le 0) { return $null }
        return $powerMilliwatts
    }
    catch {
        return $null
    }
}

function Get-ProcessCpuUtilization {
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
        return $null
    }
}

function Update-ProcessEnergyData {
    param([int]$IntervalSeconds)
    
    $cpuData = Get-ProcessCpuUtilization
    $cpuPowerMilliwatts = Get-CpuPowerConsumption
    $systemPowerMilliwatts = Get-SystemPowerConsumption
    
    if ($null -eq $cpuData -or $null -eq $cpuPowerMilliwatts) {
        return $false
    }
    
    # Update global power readings
    $script:CurrentCpuPower = $cpuPowerMilliwatts
    $script:CurrentSystemPower = if ($null -ne $systemPowerMilliwatts) { $systemPowerMilliwatts } else { 0 }
    $script:CurrentTotalCpuPercent = $cpuData.TotalCpu
    $script:MeasurementCount++
    
    $totalCpu = $cpuData.TotalCpu
    
    if ($totalCpu -gt 0) {
        foreach ($process in $cpuData.ProcessData.GetEnumerator()) {
            $processName = $process.Key
            $processCpu = $process.Value
            $cpuRatio = $processCpu / $totalCpu
            
            $processPowerCpu = $cpuPowerMilliwatts * $cpuRatio
            $energyCpu = $processPowerCpu * $IntervalSeconds
            
            $energySystem = 0
            if ($null -ne $systemPowerMilliwatts) {
                $processPowerSystem = $systemPowerMilliwatts * $cpuRatio
                $energySystem = $processPowerSystem * $IntervalSeconds
            }
            
            if (-not $script:ProcessEnergyData.ContainsKey($processName)) {
                $script:ProcessEnergyData[$processName] = @{
                    CpuEnergy = 0
                    SystemEnergy = 0
                    LastSeenCpu = 0
                    LastSeenPowerMw = 0
                    PowerHistory = New-Object System.Collections.ArrayList
                }
            }
            
            $script:ProcessEnergyData[$processName].CpuEnergy += $energyCpu
            $script:ProcessEnergyData[$processName].SystemEnergy += $energySystem
            $script:ProcessEnergyData[$processName].LastSeenCpu = $processCpu
            $script:ProcessEnergyData[$processName].LastSeenPowerMw = $processPowerCpu
            
            # Add to power history
            if ($script:ProcessEnergyData[$processName].PowerHistory.Count -ge $script:MaxHistorySize) {
                $script:ProcessEnergyData[$processName].PowerHistory.RemoveAt(0)
            }
            $null = $script:ProcessEnergyData[$processName].PowerHistory.Add($processPowerCpu)
        }
    }
    
    return $true
}

#endregion
# regions are cool :), hope u can see them
# yeah I see them, didn't even know it was a thing (〃￣︶￣)人(￣︶￣〃)