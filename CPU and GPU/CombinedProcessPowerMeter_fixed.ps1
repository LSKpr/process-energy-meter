<#
.SYNOPSIS
    Interactive Process Power Monitor
.DESCRIPTION
    Command-line tool to monitor process power consumption with:
    - Top X processes view
    - Detailed single process monitoring with power graph
    - Real-time power consumption tracking
    
    Requires: Administrator privileges and LibreHardwareMonitorLib.dll
#>

param(
    [Nullable[int]]$Duration = $null,
    [int]$SampleInterval = 100,
    [double]$WeightSM = 1.0,
    [double]$WeightMem = 0.5,
    [double]$WeightEnc = 0.25,
    [double]$WeightDec = 0.15,
    [string]$DiagnosticsOutput = $null
)

# Color scheme
$script:Colors = @{
    Header = 'Cyan'
    ProcessName = 'Yellow'
    Value = 'Green'
    Warning = 'Red'
    Info = 'White'
    Highlight = 'Magenta'
    Graph = 'Blue'
}

$script:Computer = $null
$script:CpuHardware = $null

#region Hardware Monitoring Functions

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

#endregion


# regions are cool :), hope u can see them
# yeah I see them, didn't even know it was a thing (〃￣︶￣)人(￣︶￣〃)
#region Utility Functions

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

function Draw-PowerGraph {
    param(
        [array]$PowerHistory,
        [int]$Width = 60,
        [int]$Height = 10
    )
    
    if ($PowerHistory.Count -eq 0) {
        Write-Host "  No data yet..." -ForegroundColor $script:Colors.Info
        return
    }
    
    $maxPower = ($PowerHistory | Measure-Object -Maximum).Maximum
    if ($maxPower -eq 0) { $maxPower = 1 }
    
    Write-Host "`n  Power History (Last $($PowerHistory.Count) measurements):" -ForegroundColor $script:Colors.Info
    Write-Host ("  Max: {0:N3} W" -f ($maxPower / 1000)) -ForegroundColor $script:Colors.Value
    
    # Draw graph from top to bottom
    for ($row = $Height; $row -gt 0; $row--) {
        $threshold = ($maxPower / $Height) * $row
        $line = "  "
        
        foreach ($power in $PowerHistory | Select-Object -Last $Width) {
            if ($power -ge $threshold) {
                $line += "#"
            }
            else {
                $line += " "
            }
        }
        
        Write-Host $line -ForegroundColor $script:Colors.Graph
    }
    
    # Draw baseline
    Write-Host ("  " + ("-" * ([Math]::Min($PowerHistory.Count, $Width)))) -ForegroundColor $script:Colors.Graph
    Write-Host ("  Min: {0:N3} W" -f (($PowerHistory | Measure-Object -Minimum).Minimum / 1000)) -ForegroundColor $script:Colors.Value
}

#endregion



#region Command Line Interface

$script:CurrentTopList = @()
$script:ProcessEnergyData = @{}
$script:CurrentView = "list"
$script:CurrentViewParam = 20
$script:MeasurementHistory = New-Object System.Collections.ArrayList
$script:MaxHistorySize = 100
$script:MonitoringStartTime = Get-Date
$script:MeasurementCount = 0
$script:CurrentCpuPower = 0
$script:CurrentSystemPower = 0
$script:CurrentTotalCpuPercent = 0
$script:MeasurementInterval = 2

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

function Show-TopList {
    param(
        [int]$Count,
        [bool]$ClearScreen = $true
    )
    
    if ($Count -lt 1) { $Count = 1 }
    if ($Count -gt 50) { $Count = 50 }
    
    $elapsed = (Get-Date) - $script:MonitoringStartTime
    
    # Calculate total energy
    $totalCpuEnergy = 0
    $totalSystemEnergy = 0
    foreach ($value in $script:ProcessEnergyData.Values) {
        $totalCpuEnergy += $value.CpuEnergy
        $totalSystemEnergy += $value.SystemEnergy
    }
    
    if ($ClearScreen) {
        Clear-Host
    } else {
        # Move to top-left without clearing
        [Console]::SetCursorPosition(0, 0)
    }
    
    Write-Host "`n========================================================" -ForegroundColor $script:Colors.Header
    Write-Host "   Process Power Monitor - Top $Count                    " -ForegroundColor $script:Colors.Header
    Write-Host "========================================================" -ForegroundColor $script:Colors.Header
    Write-Host ""
    
    # System stats
    Write-Host "System Statistics:" -ForegroundColor $script:Colors.Info
    Write-Host ("  Runtime:             {0:hh\:mm\:ss}" -f $elapsed) -ForegroundColor $script:Colors.Value
    Write-Host ("  Measurements:        {0}" -f $script:MeasurementCount) -ForegroundColor $script:Colors.Value
    Write-Host ("  Measurement Interval: {0}s" -f $script:MeasurementInterval) -ForegroundColor $script:Colors.Value
    Write-Host ("  Tracked Processes:   {0}" -f $script:ProcessEnergyData.Count) -ForegroundColor $script:Colors.Value
    Write-Host ""
    Write-Host ("  Total CPU Usage:     {0:N1}%" -f $script:CurrentTotalCpuPercent) -ForegroundColor $script:Colors.Value
    Write-Host ("  CPU Power Now:       {0:N2} W" -f ($script:CurrentCpuPower / 1000)) -ForegroundColor $script:Colors.Value
    if ($script:CurrentSystemPower -gt 0) {
        Write-Host ("  System Power Now:    {0:N2} W" -f ($script:CurrentSystemPower / 1000)) -ForegroundColor $script:Colors.Value
    }
    Write-Host ("  Total CPU Energy:    {0}" -f (Format-EnergyValue $totalCpuEnergy)) -ForegroundColor Yellow
    if ($totalSystemEnergy -gt 0) {
        Write-Host ("  Total System Energy: {0}" -f (Format-EnergyValue $totalSystemEnergy)) -ForegroundColor Yellow
    }
    Write-Host ""
    
    Write-Host "Top $Count Processes:" -ForegroundColor $script:Colors.Highlight
    Write-Host ("{0,-5} {1,-30} {2,12} {3,10} {4,12}" -f "#", "Process", "CPU Energy", "CPU %", "Power Now") -ForegroundColor $script:Colors.Header
    Write-Host ("-" * 75) -ForegroundColor $script:Colors.Header
    
    $topProcesses = $script:ProcessEnergyData.GetEnumerator() | 
        Sort-Object { $_.Value.CpuEnergy } -Descending |
        Select-Object -First $Count
    
    $script:CurrentTopList = @()
    $rank = 1
    
    foreach ($proc in $topProcesses) {
        $script:CurrentTopList += $proc.Key
        
        $cpuEnergy = Format-EnergyValue $proc.Value.CpuEnergy
        $currentCpu = "{0:N2}%" -f $proc.Value.LastSeenCpu
        $currentPower = "{0:N3} W" -f ($proc.Value.LastSeenPowerMw / 1000)
        
        $color = switch ($rank) {
            1 { 'Yellow' }
            2 { 'Green' }
            3 { 'Cyan' }
            default { 'White' }
        }
        
        Write-Host ("{0,-5} {1,-30} {2,12} {3,10} {4,12}" -f $rank, $proc.Key, $cpuEnergy, $currentCpu, $currentPower) -ForegroundColor $color
        $rank++
    }
    
    Write-Host ""
    Write-Host "Commands: list X | focus X | interval X | help | quit" -ForegroundColor $script:Colors.Info
    
    # Clear remaining lines and position cursor for input
    if (-not $ClearScreen) {
        $currentLine = [Console]::CursorTop
        $windowHeight = [Console]::WindowHeight
        for ($i = $currentLine; $i -lt ($windowHeight - 2); $i++) {
            Write-Host (" " * [Console]::WindowWidth)
        }
        # Move cursor to input line at bottom
        [Console]::SetCursorPosition(0, $windowHeight - 2)
    }
    
    Write-Host "Command> " -NoNewline -ForegroundColor $script:Colors.Highlight
}

function Show-FocusedView {
    param([string]$ProcessName)
    
    if (-not $script:ProcessEnergyData.ContainsKey($ProcessName)) {
        Clear-Host
        Write-Host "`nProcess '$ProcessName' not found or has no data yet." -ForegroundColor $script:Colors.Warning
        Write-Host "Press any key to return..." -ForegroundColor $script:Colors.Info
        $null = [Console]::ReadKey($true)
        return
    }
    
    $processData = $script:ProcessEnergyData[$ProcessName]
    $elapsed = (Get-Date) - $script:MonitoringStartTime
    
    Clear-Host
    Write-Host "`n========================================================" -ForegroundColor $script:Colors.Header
    Write-Host "   Focused View: $ProcessName" -ForegroundColor $script:Colors.Header
    Write-Host "========================================================" -ForegroundColor $script:Colors.Header
    Write-Host ""
    
    Write-Host "Current Status:" -ForegroundColor $script:Colors.Info
    Write-Host ("  CPU Usage:           {0:N2}% of {1:N1}%" -f $processData.LastSeenCpu, $script:CurrentTotalCpuPercent) -ForegroundColor $script:Colors.Value
    Write-Host ("  Process Power:       {0:N3} W" -f ($processData.LastSeenPowerMw / 1000)) -ForegroundColor $script:Colors.Value
    Write-Host ("  CPU Power Now:       {0:N2} W" -f ($script:CurrentCpuPower / 1000)) -ForegroundColor $script:Colors.Value
    if ($script:CurrentSystemPower -gt 0) {
        Write-Host ("  System Power Now:    {0:N2} W" -f ($script:CurrentSystemPower / 1000)) -ForegroundColor $script:Colors.Value
    } else {
        Write-Host "  System Power Now:    N/A" -ForegroundColor Yellow
    }
    Write-Host ""
    
    Write-Host "Accumulated (since program start):" -ForegroundColor $script:Colors.Info
    Write-Host ("  CPU Energy:          {0}" -f (Format-EnergyValue $processData.CpuEnergy)) -ForegroundColor $script:Colors.Value
    if ($processData.SystemEnergy -gt 0) {
        Write-Host ("  System Energy:       {0}" -f (Format-EnergyValue $processData.SystemEnergy)) -ForegroundColor $script:Colors.Value
    }
    Write-Host ("  Runtime:             {0:hh\:mm\:ss}" -f $elapsed) -ForegroundColor $script:Colors.Value
    
    $avgPowerW = if ($elapsed.TotalSeconds -gt 0) { ($processData.CpuEnergy / 1000) / $elapsed.TotalSeconds } else { 0 }
    Write-Host ("  Average Power:       {0:N3} W" -f $avgPowerW) -ForegroundColor Yellow
    
    # Draw power graph
    if ($processData.PowerHistory.Count -gt 0) {
        Draw-PowerGraph -PowerHistory $processData.PowerHistory -Width 60 -Height 8
    }
    
    Write-Host ""
    Write-Host "Commands: list X | focus X | interval X | help | quit" -ForegroundColor $script:Colors.Info
    Write-Host "Command> " -NoNewline -ForegroundColor $script:Colors.Highlight
}

function Show-CommandHelp {
    Clear-Host
    Write-Host "`n========================================================" -ForegroundColor $script:Colors.Header
    Write-Host "   Command Help                                         " -ForegroundColor $script:Colors.Header
    Write-Host "========================================================" -ForegroundColor $script:Colors.Header
    Write-Host ""
    Write-Host "Available Commands:" -ForegroundColor $script:Colors.Highlight
    Write-Host ""
    Write-Host "  list X      - Show top X processes by energy consumption" -ForegroundColor $script:Colors.Info
    Write-Host "                Example: list 5, list 10, list 20" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  focus X     - Show detailed view of process #X from list" -ForegroundColor $script:Colors.Info
    Write-Host "                Example: focus 1 (focuses on #1 process)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  interval X  - Change measurement interval (1-60 seconds)" -ForegroundColor $script:Colors.Info
    Write-Host "                Example: interval 1, interval 5" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  help        - Show this help message" -ForegroundColor $script:Colors.Info
    Write-Host ""
    Write-Host "  quit/exit   - Exit the program" -ForegroundColor $script:Colors.Info
    Write-Host ""
    Write-Host "Note: All views auto-update. Monitoring runs continuously" -ForegroundColor Yellow
    Write-Host "      in background. Energy is accumulated from program start." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Press any key to return..." -ForegroundColor $script:Colors.Info
    $null = [Console]::ReadKey($true)
}

function Start-CommandLineMode {
    param([int]$IntervalSeconds)
    
    Clear-Host
    Write-Host "`n========================================================" -ForegroundColor $script:Colors.Header
    Write-Host "   Interactive Process Power Monitor                    " -ForegroundColor $script:Colors.Header
    Write-Host "========================================================" -ForegroundColor $script:Colors.Header
    Write-Host ""
    Write-Host "CPU: $($script:CpuHardware.Name)" -ForegroundColor $script:Colors.Value
    
    $testSystemPower = Get-SystemPowerConsumption
    if ($null -eq $testSystemPower) {
        Write-Host "System Power: Not Available" -ForegroundColor Yellow
    }
    else {
        Write-Host "System Power: Available" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "Collecting initial data..." -ForegroundColor $script:Colors.Info
    
    # Collect initial data
    for ($i = 0; $i -lt 3; $i++) {
        $null = Update-ProcessEnergyData -IntervalSeconds $IntervalSeconds
        Start-Sleep -Seconds $IntervalSeconds
        Write-Host "." -NoNewline -ForegroundColor $script:Colors.Info
    }
    
    Write-Host "`n`nStarting with list 20 view...`n" -ForegroundColor Green
    Start-Sleep -Seconds 1
    
    # Set initial view
    $script:CurrentView = "list"
    $script:CurrentViewParam = 20
    $script:MonitoringStartTime = Get-Date
    
    # Hide cursor for cleaner display
    [Console]::CursorVisible = $false
    
    # Set initial interval
    $script:MeasurementInterval = $IntervalSeconds
    
    # Track last update times
    $lastDataUpdate = Get-Date
    $lastDisplayRefresh = Get-Date
    
    try {
        # Show initial display
        Show-TopList -Count $script:CurrentViewParam -ClearScreen $true
        [Console]::CursorVisible = $true
        
        # Main input loop
        while ($true) {
            $now = Get-Date
            
            # Update data based on current interval
            if (($now - $lastDataUpdate).TotalSeconds -ge $script:MeasurementInterval) {
                # Use ACTUAL elapsed time for accurate energy calculation
                $actualInterval = ($now - $lastDataUpdate).TotalSeconds
                $null = Update-ProcessEnergyData -IntervalSeconds $actualInterval
                $lastDataUpdate = $now
            }
            
            # Refresh display every 2 seconds
            if (($now - $lastDisplayRefresh).TotalSeconds -ge 2) {
                # Save cursor position
                $cursorLeft = [Console]::CursorLeft
                $cursorTop = [Console]::CursorTop
                
                # Update display without clearing input line
                [Console]::CursorVisible = $false
                switch ($script:CurrentView) {
                    "list" { Show-TopList -Count $script:CurrentViewParam -ClearScreen $false }
                    "focus" { Show-FocusedView -ProcessName $script:CurrentViewParam }
                }
                
                # Restore cursor position for input
                [Console]::SetCursorPosition($cursorLeft, [Console]::WindowHeight - 2)
                [Console]::CursorVisible = $true
                $lastDisplayRefresh = $now
            }
            
            # Check for user input (non-blocking)
            if ([Console]::KeyAvailable) {
                [Console]::CursorVisible = $true
                
                # Read the entire line using native ReadLine
                $inputLine = [Console]::ReadLine()
                
                if (-not [string]::IsNullOrWhiteSpace($inputLine)) {
                    $parts = $inputLine.Trim() -split '\s+'
                    $cmd = $parts[0].ToLower()
                    
                    switch ($cmd) {
                        'list' {
                            $count = 20
                            if ($parts.Length -gt 1) {
                                [int]::TryParse($parts[1], [ref]$count) | Out-Null
                            }
                            $script:CurrentView = "list"
                            $script:CurrentViewParam = $count
                            Show-TopList -Count $script:CurrentViewParam -ClearScreen $true
                            [Console]::CursorVisible = $true
                        }
                        'focus' {
                            if ($parts.Length -lt 2) {
                                Write-Host "`nUsage: focus X (where X is process number)" -ForegroundColor $script:Colors.Warning
                                Start-Sleep -Seconds 2
                                Show-TopList -Count $script:CurrentViewParam -ClearScreen $true
                                [Console]::CursorVisible = $true
                            }
                            else {
                                $index = 0
                                if ([int]::TryParse($parts[1], [ref]$index) -and $index -gt 0 -and $index -le $script:CurrentTopList.Count) {
                                    $processName = $script:CurrentTopList[$index - 1]
                                    $script:CurrentView = "focus"
                                    $script:CurrentViewParam = $processName
                                    Show-FocusedView -ProcessName $script:CurrentViewParam
                                    [Console]::CursorVisible = $true
                                }
                                else {
                                    Write-Host "`nInvalid process number" -ForegroundColor $script:Colors.Warning
                                    Start-Sleep -Seconds 2
                                    Show-TopList -Count $script:CurrentViewParam -ClearScreen $true
                                    [Console]::CursorVisible = $true
                                }
                            }
                        }
                        'interval' {
                            if ($parts.Length -lt 2) {
                                Write-Host "`nUsage: interval X (1-60 seconds)" -ForegroundColor $script:Colors.Warning
                                Start-Sleep -Seconds 2
                            }
                            else {
                                $newInterval = 0
                                if ([int]::TryParse($parts[1], [ref]$newInterval) -and $newInterval -ge 1 -and $newInterval -le 60) {
                                    $script:MeasurementInterval = $newInterval
                                    Write-Host "`nMeasurement interval set to $newInterval seconds" -ForegroundColor Green
                                    Start-Sleep -Seconds 1
                                }
                                else {
                                    Write-Host "`nInvalid interval. Use 1-60 seconds." -ForegroundColor $script:Colors.Warning
                                    Start-Sleep -Seconds 2
                                }
                            }
                            Show-TopList -Count $script:CurrentViewParam -ClearScreen $true
                            [Console]::CursorVisible = $true
                        }
                        'help' {
                            Show-CommandHelp
                            [Console]::CursorVisible = $true
                        }
                        'quit' {
                            throw "EXIT"
                        }
                        'exit' {
                            throw "EXIT"
                        }
                        default {
                            Write-Host "`nIncorrect command. Type 'help' for available commands." -ForegroundColor $script:Colors.Warning
                            Start-Sleep -Seconds 2
                            # Refresh current view properly
                            switch ($script:CurrentView) {
                                "list" { Show-TopList -Count $script:CurrentViewParam -ClearScreen $true }
                                "focus" { Show-FocusedView -ProcessName $script:CurrentViewParam }
                            }
                            [Console]::CursorVisible = $true
                        }
                    }
                }
                else {
                    # Empty input - just refresh
                    Show-TopList -Count $script:CurrentViewParam -ClearScreen $true
                    [Console]::CursorVisible = $true
                }
            }
            
            Start-Sleep -Milliseconds 100
        }
    }
    finally {
        # Cleanup
        [Console]::CursorVisible = $true
        
        if ($_.Exception.Message -eq "EXIT") {
            Write-Host "`nExiting..." -ForegroundColor $script:Colors.Info
        }
    }
}

#endregion

#region Main Execution

# Check admin privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "`nERROR: This script requires Administrator privileges!" -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again.`n" -ForegroundColor Yellow
    exit 1
}

Clear-Host
Write-Host "`nInitializing Interactive Process Power Monitor..." -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor Cyan

$initialized = Initialize-LibreHardwareMonitor

if (-not $initialized) {
    Write-Host "`nFailed to initialize. Exiting..." -ForegroundColor Red
    exit 1
}

Write-Host "[OK] Hardware monitoring ready" -ForegroundColor Green

Start-CommandLineMode -IntervalSeconds $Duration

# Cleanup
if ($null -ne $script:Computer) {
    $script:Computer.Close()
}

Write-Host "`nThank you for using Process Power Monitor!`n" -ForegroundColor Cyan

#endregion


<#
Flags of interest:
    -Duration: how long to monitor in seconds (default 60)
    -SampleInterval: how often to sample in ms (default 100)
    -WeightSM, WeightMem, WeightEnc, WeightDec: weights for the attribution formula (defaults: 1.0, 0.5, 0.25, 0.15)
    -DiagnosticsOutput: base path for writing diagnostics CSV files (default "power_diagnostics")
#>

function Write-Log($msg) {
    $t = Get-Date
    $tStr = $t.ToString('yyyy/MM/dd HH:mm:ss.fff')
    Write-Host "[$tStr] $msg"
}

class GPUProcessMonitor {
    [int]$SampleIntervalMs
    [double]$WeightSm
    [double]$WeightMemory
    [double]$WeightEncoder
    [double]$WeightDecoder
    [System.Collections.Generic.List[object]]$Samples
    [System.Diagnostics.Stopwatch]$Stopwatch
    [int64]$StopwatchFrequency
    [long]$LastSampleTime
    [datetime]$TimeStampLogging
    [hashtable]$ProcessEnergyJoules
    [string]$GpuName
    [double]$GpuEnergyJoules
    [double]$GpuIdlePower
    [double]$GpuIdlePowerMin
    [double]$GpuIdlePowerMax
    [double]$GpuIdleFanPercent
    [double]$GpuIdleTemperatureC
    [hashtable]$GpuIdleProcesses
    [int]$IdleProcessCount
    [string]$DiagnosticsOutputPath
    [System.Collections.Generic.Dictionary[int,string]] $ProcessNameCache
    [System.Collections.Generic.LinkedList[int]] $ProcessNameCacheOrder
    [int] $ProcessNameCacheCapacity
    [System.IO.StreamWriter] $SamplesCsvWriter
    [System.IO.StreamWriter] $ProcessesCsvWriter
    [int] $SamplesSinceLastFlush
    [int] $SamplesFlushInterval
    [int] $MaxSamplesInMemory

    GPUProcessMonitor([int]$SampleIntervalMs, [double]$wSM, [double]$wMem, [double]$wEnc, [double]$wDec, [string]$diagPath) {
        $this.SampleIntervalMs = $SampleIntervalMs
        $this.WeightSm = $wSM
        $this.WeightMemory = $wMem
        $this.WeightEncoder = $wEnc
        $this.WeightDecoder = $wDec
        $this.Samples = [System.Collections.Generic.List[object]]::new()
        $this.Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $this.StopwatchFrequency = [System.Diagnostics.Stopwatch]::Frequency
        $this.LastSampleTime = $this.Stopwatch.ElapsedTicks
        $this.TimeStampLogging = Get-Date
        $this.ProcessEnergyJoules = @{}
        $this.GpuEnergyJoules = 0.0
        $this.GpuIdleProcesses = @{}
        $this.DiagnosticsOutputPath = $diagPath

        # initialize LRU cache defaults
        $this.ProcessNameCache = $null
        $this.ProcessNameCacheOrder = $null
        $this.ProcessNameCacheCapacity = 1024       # default capacity

        # writer flush tuning
        $this.SamplesSinceLastFlush = 0
        $this.SamplesFlushInterval = 10   # flush every 10 samples

        # long-run safety: cap in-memory samples
        $this.MaxSamplesInMemory = 10000  # default; adjust if you want to keep more samples in memory

        # Get GPU name
        $gpuNameOutput = nvidia-smi --query-gpu=name --format=csv,noheader 2>$null
        $this.GpuName = if ($gpuNameOutput) { $gpuNameOutput.Trim() } else { "unknown" }
        Write-Log("Monitoring GPU: $($this.GpuName)")
        Write-Host "Power attribution weights: SM=$wSM, Mem=$wMem, Enc=$wEnc, Dec=$wDec"

        # Measure idle metrics
        Write-Log("Measuring idle power (please ensure GPU is idle; close heavy apps).")
        $this.MeasureIdleMetrics()
        Write-Log(("Idle GPU power measured: {0:N2}W [Min: {1:N2}W  Max:{2:N2}W] with {3} Processes; Temp: {4:F1}C  Fan: {5:F1}%" -f `
            $this.GpuIdlePower, $this.GpuIdlePowerMin, $this.GpuIdlePowerMax, $this.IdleProcessCount, $this.GpuIdleTemperatureC, $this.GpuIdleFanPercent))

        # Setup diagnostics paths and open writers immediately (header creation happens here)
        try {
            $timestampForFile = $this.TimeStampLogging.ToString('yyyyMMdd_HHmmss')
            $outSpec = $this.DiagnosticsOutputPath
            if (-not $outSpec) {
                $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
                $outSpec = Join-Path $scriptDir 'power_diagnostics'
            }

            if ($outSpec -match '\.csv$') {
                $csvPath = $outSpec
                $csvPathProcesses = $outSpec -replace '\.csv$','_processes.csv'
                $diagDir = Split-Path -Parent $csvPath
                if ($diagDir -and -not (Test-Path $diagDir)) { New-Item -ItemType Directory -Path $diagDir -Force | Out-Null }
            } else {
                $diagDir = $outSpec
                if (-not (Test-Path $diagDir)) { New-Item -ItemType Directory -Path $diagDir -Force | Out-Null }
                $csvPath = Join-Path $diagDir ("samples_$timestampForFile.csv")
                $csvPathProcesses = Join-Path $diagDir ("processes_$timestampForFile.csv")
            }

            # Use FileStream with FileShare.Read so other tools can read while we append (robust for long runs)
            $needHeaderSamples = -not (Test-Path $csvPath)
            $fsSamples = [System.IO.FileStream]::new($csvPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
            # Seek to end for append
            $fsSamples.Seek(0, [System.IO.SeekOrigin]::End) | Out-Null
            $this.SamplesCsvWriter = [System.IO.StreamWriter]::new($fsSamples, [System.Text.Encoding]::UTF8)
            $this.SamplesCsvWriter.AutoFlush = $false
            if ($needHeaderSamples -and ($fsSamples.Length -eq 0)) {
                $this.SamplesCsvWriter.WriteLine('Timestamp,PowerW,ActivePowerW,ExcessPowerW,GpuSMUtil,GpuMemUtil,GpuEncUtil,GpuDecUtil,GpuWeightTotal,ProcessWeightTotal,ProcessCount,AttributedPowerW,ResidualPowerW,AccumulatedEnergyJ,TemperatureC,FanPercent')
                $this.SamplesCsvWriter.Flush()
            }

            $needHeaderProcs = -not (Test-Path $csvPathProcesses)
            $fsProcs = [System.IO.FileStream]::new($csvPathProcesses, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
            $fsProcs.Seek(0, [System.IO.SeekOrigin]::End) | Out-Null
            $this.ProcessesCsvWriter = [System.IO.StreamWriter]::new($fsProcs, [System.Text.Encoding]::UTF8)
            $this.ProcessesCsvWriter.AutoFlush = $false
            if ($needHeaderProcs -and ($fsProcs.Length -eq 0)) {
                $this.ProcessesCsvWriter.WriteLine('Timestamp,PID,ProcessName,SMUtil,MemUtil,EncUtil,DecUtil,PowerW,EnergyJ,AccumulatedEnergyJ,WeightedUtil,IsIdle')
                $this.ProcessesCsvWriter.Flush()
            }
        } catch {
            Write-Log("Warning: Could not create diagnostics files in constructor: $($_.Exception.Message)")
            # If we couldn't open writers here, Sample() will attempt lazy-open -- keep behavior safe
            $this.SamplesCsvWriter = $null
            $this.ProcessesCsvWriter = $null
        }
    }

    [void] MeasureIdleMetrics() {
        $idleTemperatureSamples = @()
        $idleFanUtilSamples = @()
        $idlePowerSamples = @()
        $idlePowerMin = 0
        $idlePowerMax = 0
        for ($loopIndex = 0; $loopIndex -lt 50; $loopIndex++) {
            $nvidiaSmiOutput = nvidia-smi --query-gpu=power.draw.instant,temperature.gpu,fan.speed --format=csv,noheader,nounits 2>$null
            if ($nvidiaSmiOutput -and -not [string]::IsNullOrWhiteSpace($nvidiaSmiOutput) -and $nvidiaSmiOutput.Trim() -ne '[N/A]') {
                $parts = $nvidiaSmiOutput -split ',\s*'
                if ($parts.Count -ge 3) {
                    try {
                        $idlePowerSamples += [double]$parts[0].Trim()
                        $idleTemperatureSamples += [double]$parts[1].Trim()
                        $idleFanUtilSamples += [double]$parts[2].Trim()
                    }
                    catch {}
                }
            }
            Start-Sleep -Milliseconds $this.SampleIntervalMs
        }

        if ($idleTemperatureSamples.Count -gt 0) { $this.GpuIdleTemperatureC = ($idleTemperatureSamples | Measure-Object -Average).Average }
        if ($idleFanUtilSamples.Count -gt 0) { $this.GpuIdleFanPercent = ($idleFanUtilSamples | Measure-Object -Average).Average }
        if ($idlePowerSamples.Count -gt 0) {
            $measuredIdlePower = ($idlePowerSamples | Measure-Object -Average).Average
            $idlePowerMin = $idlePowerSamples | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
            $idlePowerMax = $idlePowerSamples | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
            $this.GpuIdlePower = $measuredIdlePower
            $this.GpuIdlePowerMin = $idlePowerMin
            $this.GpuIdlePowerMax = $idlePowerMax
        }

        # Record which processes were running during idle measurement
        $idleProcesses = $this.GetGpuProcesses()
        $this.IdleProcessCount = $idleProcesses.Count
        foreach ($idleProcessObj in $idleProcesses) {
            $this.GpuIdleProcesses[[string]$idleProcessObj.ProcessId] = $true
        }
    }

    # ---------- optimized GetGpuProcesses() --------------------------------
    [array] GetGpuProcesses() {
        if ($null -eq $this.ProcessNameCache) {
            $this.ProcessNameCache = [System.Collections.Generic.Dictionary[int,string]]::new()
            $this.ProcessNameCacheOrder = [System.Collections.Generic.LinkedList[int]]::new()
            if ($this.ProcessNameCacheCapacity -le 0) { $this.ProcessNameCacheCapacity = 1024 }
        }

        $resultList = [System.Collections.Generic.List[object]]::new()

        $pmonOutput = nvidia-smi pmon -c 1 2>$null
        if (-not $pmonOutput) { return $resultList.ToArray() }

        $pmonPattern = '^\s*(\d+)\s+(\d+)\s+([A-Z+]+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(.+?)\s*$'
        $pmonRegex = [System.Text.RegularExpressions.Regex]::new($pmonPattern, [System.Text.RegularExpressions.RegexOptions]::Compiled)

        $rawEntries = [System.Collections.Generic.List[object]]::new()
        $pidHash = [System.Collections.Generic.HashSet[int]]::new()

        foreach ($line in ($pmonOutput -split "`n")) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if ($line[0] -eq '#') { continue }

            $regexMatch = $pmonRegex.Match($line)
            if (-not $regexMatch.Success) { continue }

            $procIdInt = [int]$regexMatch.Groups[2].Value

            [void]$pidHash.Add($procIdInt)

            $rawEntries.Add([PSCustomObject]@{
                ProcessId = $procIdInt
                SmUtil    = $this.ParseUtil($regexMatch.Groups[4].Value)
                MemUtil   = $this.ParseUtil($regexMatch.Groups[5].Value)
                EncUtil   = $this.ParseUtil($regexMatch.Groups[6].Value)
                DecUtil   = $this.ParseUtil($regexMatch.Groups[7].Value)
                Command   = $regexMatch.Groups[10].Value.Trim()
            })
        }

        if ($rawEntries.Count -eq 0) { return $resultList.ToArray() }

        # local references for performance
        $processNameCacheLocal = $this.ProcessNameCache
        $processNameCacheOrderLocal = $this.ProcessNameCacheOrder
        $processNameCacheCapacityLocal = $this.ProcessNameCacheCapacity

        # Build list of PIDs missing from cache
        $pidsToResolve = [System.Collections.Generic.List[int]]::new()
        foreach ($pidCandidate in $pidHash) {
            if (-not $processNameCacheLocal.ContainsKey($pidCandidate)) {
                $pidsToResolve.Add($pidCandidate)
            } else {
                # update LRU position (move to front)
                try {
                    $node = $processNameCacheOrderLocal.Find($pidCandidate)
                    if ($node) {
                        $processNameCacheOrderLocal.Remove($node)
                        $processNameCacheOrderLocal.AddFirst($node)
                    }
                } catch {}
            }
        }

        if ($pidsToResolve.Count -gt 0) {
            try {
                $pidArray = $pidsToResolve.ToArray()
                $winProcs = Get-Process -Id $pidArray -ErrorAction SilentlyContinue
                foreach ($winProc in $winProcs) {
                    $winProcId = [int]$winProc.Id
                    $winProcName = $winProc.ProcessName
                    if (-not $processNameCacheLocal.ContainsKey($winProcId)) {
                        $processNameCacheLocal.Add($winProcId, $winProcName)
                        $processNameCacheOrderLocal.AddFirst($winProcId)
                        while ($processNameCacheOrderLocal.Count -gt $processNameCacheCapacityLocal) {
                            $lastNode = $processNameCacheOrderLocal.Last
                            if ($lastNode) {
                                $oldPid = [int]$lastNode.Value
                                $processNameCacheOrderLocal.RemoveLast()
                                if ($processNameCacheLocal.ContainsKey($oldPid)) { $processNameCacheLocal.Remove($oldPid) }
                            } else { break }
                        }
                    } else {
                        $existingNode = $processNameCacheOrderLocal.Find($winProcId)
                        if ($existingNode) {
                            $processNameCacheOrderLocal.Remove($existingNode)
                            $processNameCacheOrderLocal.AddFirst($existingNode)
                        }
                    }
                }
            } catch {
                # ignore name resolution failures (best-effort)
            }
        }

        # Build final list (use cached name if present, otherwise fallback to pmon Command)
        foreach ($entry in $rawEntries) {
            $resolved = $entry.Command
            $pidKey = [int]$entry.ProcessId
            $tmpName = $null
            if ($processNameCacheLocal.TryGetValue($pidKey, [ref]$tmpName)) {
                $resolved = $tmpName
                try {
                    $node = $processNameCacheOrderLocal.Find($pidKey)
                    if ($node) {
                        $processNameCacheOrderLocal.Remove($node)
                        $processNameCacheOrderLocal.AddFirst($node)
                    }
                } catch {}
            }

            $resultList.Add([PSCustomObject]@{
                ProcessId    = $entry.ProcessId
                ProcessName  = [string]$resolved
                SmUtil       = $entry.SmUtil
                MemUtil      = $entry.MemUtil
                EncUtil      = $entry.EncUtil
                DecUtil      = $entry.DecUtil
                WeightedUtil = 0.0
            })
        }

        return $resultList.ToArray()
    }

    [double] ParseUtil([string]$value) {
        if ([string]::IsNullOrWhiteSpace($value) -or $value -eq '-') { return 0.0 }

        $trimmedString = $value.Trim()
        $trimmedString = $trimmedString -replace '[^\d\.\-+]',''

        $parsedDouble = 0.0
        if ([double]::TryParse($trimmedString, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedDouble)) {
            return $parsedDouble
        }

        return 0.0
    }

    # helper to safely flush+close writers (called on exit)
    [void] CloseWriters() {
        try {
            if ($null -ne $this.SamplesCsvWriter) {
                try { $this.SamplesCsvWriter.Flush() } catch {}
                try { $this.SamplesCsvWriter.Close() } catch {}
                $this.SamplesCsvWriter = $null
            }
        } catch {}

        try {
            if ($null -ne $this.ProcessesCsvWriter) {
                try { $this.ProcessesCsvWriter.Flush() } catch {}
                try { $this.ProcessesCsvWriter.Close() } catch {}
                $this.ProcessesCsvWriter = $null
            }
        } catch {}
    }

    # ---------- optimized Sample() ----------------------------------------
    [void] Sample() {
        # fast local caches of instance fields (avoid repeated $this. property lookup)
        $weightSmLocal      = $this.WeightSm
        $weightMemoryLocal  = $this.WeightMemory
        $weightEncoderLocal = $this.WeightEncoder
        $weightDecoderLocal = $this.WeightDecoder
        $gpuIdleProcessMap  = $this.GpuIdleProcesses
        $processEnergyMap   = $this.ProcessEnergyJoules

        # timing (hot path)
        $currentTicks = $this.Stopwatch.ElapsedTicks
        $deltaTicks   = $currentTicks - $this.LastSampleTime
        $this.LastSampleTime = $currentTicks
        if ($deltaTicks -le 0) { return }
        $deltaSeconds = $deltaTicks / $this.StopwatchFrequency

        # GPU-wide metrics (batched)
        $timestamp = "-"
        $gpuPower = 0.0
        $gpuSmUtil = 0.0
        $gpuMemUtil = 0.0
        $gpuEncUtil = 0.0
        $gpuDecUtil = 0.0
        $gpuTemperature = 0.0
        $gpuFanUtil = 0.0

        # Acquire GPU metrics and per-process entries
        $combinedOut = nvidia-smi --query-gpu=timestamp,power.draw.instant,utilization.gpu,utilization.memory,utilization.encoder,utilization.decoder,temperature.gpu,fan.speed --format=csv,noheader,nounits 2>$null
        $processEntries = $this.GetGpuProcesses()
        $parts = $combinedOut -split ',\s*'
        if ($parts.Count -ge 8) {
            try {
                $timestamp = $parts[0]
                $gpuPower = [double]$parts[1].Trim()
                $gpuSmUtil = [double]$parts[2].Trim()
                $gpuMemUtil = [double]$parts[3].Trim()
                $gpuEncUtil = [double]$parts[4].Trim()
                $gpuDecUtil = [double]$parts[5].Trim()
                $gpuTemperature = [double]$parts[6].Trim()
                $gpuFanUtil = [double]$parts[7].Trim()
            } catch {}
        }
        $currentProcessCount = $processEntries.Count

        # weighted GPU total
        $gpuWeightedTotal = ($weightSmLocal * $gpuSmUtil) +
                            ($weightMemoryLocal * $gpuMemUtil) +
                            ($weightEncoderLocal * $gpuEncUtil) +
                            ($weightDecoderLocal * $gpuDecUtil)

        # per-process weighted totals
        $processWeightTotal = 0.0
        $inactiveIdleProcessesCount = 0
        foreach ($processEntry in $processEntries) {
            $weightedValue = ($weightSmLocal * $processEntry.SmUtil) +
                            ($weightMemoryLocal * $processEntry.MemUtil) +
                            ($weightEncoderLocal * $processEntry.EncUtil) +
                            ($weightDecoderLocal * $processEntry.DecUtil)
            $processEntry.WeightedUtil = $weightedValue
            $processWeightTotal += $weightedValue

            $processIdStringForLookup = $processEntry.ProcessId.ToString()
            $isIdleLookup = $gpuIdleProcessMap.ContainsKey($processIdStringForLookup)
            if ($weightedValue -le 0 -and $isIdleLookup) {
                $inactiveIdleProcessesCount++
            }
        }

        # attribution math
        $gpuActivePower = if (($this.GpuIdlePower) -lt $gpuPower) { $gpuPower - ($this.GpuIdlePower) } else { 0 }
        $gpuTotalProcessUtil = if ($gpuWeightedTotal -gt 0) { $processWeightTotal / $gpuWeightedTotal } else { 0 }
        $gpuTotalActiveProcessPower = $gpuTotalProcessUtil * $gpuActivePower
        $gpuExcessPower = $gpuActivePower - $gpuTotalActiveProcessPower

        $activeProcessCount = if ($inactiveIdleProcessesCount -eq $currentProcessCount) { $currentProcessCount } else { $currentProcessCount - $inactiveIdleProcessesCount }
        $P_idle_pwr = if ($this.IdleProcessCount -gt 0) { $this.GpuIdlePower / $this.IdleProcessCount } else { 0 }
        $P_residual_per_proc = if ($activeProcessCount -gt 0) { $gpuExcessPower / $activeProcessCount } else { 0 }

        # per-process attribution objects
        $sampleAttributedPower = 0.0
        $samplePerProcess = @{ }

        # Precompute culture once
        $invariantCulture = [System.Globalization.CultureInfo]::InvariantCulture

        # Precompute timestamp sanitized once
        $tsEsc_local = ($timestamp -replace '"','""')

        foreach ($processEntry in $processEntries) {
            $processIdString = $processEntry.ProcessId.ToString()
            $weightedForThisProcess = $processEntry.WeightedUtil

            if ($processWeightTotal -le 0) { $fraction = 0 } else { $fraction = $weightedForThisProcess / $processWeightTotal }

            $isIdleThisProcess = $gpuIdleProcessMap.ContainsKey($processIdString)
            if ($isIdleThisProcess) {
                if ($weightedForThisProcess -le 0) {
                    $powerValue = if ($inactiveIdleProcessesCount -eq $currentProcessCount) { $P_idle_pwr + $P_residual_per_proc } else { $P_idle_pwr }
                } else {
                    $powerValue = $P_idle_pwr + ($fraction * $gpuTotalActiveProcessPower) + $P_residual_per_proc
                }
            } else {
                if ($weightedForThisProcess -le 0) {
                    $powerValue = $P_residual_per_proc
                } else {
                    $powerValue = ($fraction * $gpuTotalActiveProcessPower) + $P_residual_per_proc
                }
            }

            $energyJ = $powerValue * $deltaSeconds

            # update ProcessEnergyJoules safely and quickly (strings as keys)
            if ($processEnergyMap.ContainsKey($processIdString)) {
                $existingEnergyValue = [double]$processEnergyMap[$processIdString]
            } else {
                $existingEnergyValue = 0.0
            }
            $accumulatedProcessEnergy = $existingEnergyValue + $energyJ
            $processEnergyMap[$processIdString] = $accumulatedProcessEnergy

            $sampleAttributedPower += $powerValue

            # build per-process PSCustomObject with raw numeric values (rounding only at logging)
            $samplePerProcess[$processIdString] = [PSCustomObject]@{
                PID                       = $processIdString
                ProcessName               = [string]$processEntry.ProcessName
                SMUtil                    = $processEntry.SmUtil
                MemUtil                   = $processEntry.MemUtil
                EncUtil                   = $processEntry.EncUtil
                DecUtil                   = $processEntry.DecUtil
                PowerW                    = $powerValue
                EnergyJ                   = $energyJ
                AccumulatedProcessEnergyJ = $accumulatedProcessEnergy
                WeightedUtil              = $weightedForThisProcess
                IsIdle                    = $isIdleThisProcess
            }
        }

        # record sample in memory
        $this.GpuEnergyJoules += ($gpuPower * $deltaSeconds)
        $accumulatedEnergy = $this.GpuEnergyJoules
        $sampleResidualPower = $gpuPower - $sampleAttributedPower

        # Add sample to in-memory list and enforce memory cap (avoid unbounded growth during very long runs)
        $sampleObj = [PSCustomObject]@{
            Timestamp          = $timestamp
            PowerW             = $gpuPower
            ActivePowerW       = $gpuActivePower
            ExcessPowerW       = $gpuExcessPower
            GpuSmUtil          = $gpuSmUtil
            GpuMemUtil         = $gpuMemUtil
            GpuEncUtil         = $gpuEncUtil
            GpuDecUtil         = $gpuDecUtil
            GpuWeightedTotal   = $gpuWeightedTotal
            ProcessWeightTotal = $processWeightTotal
            ProcessCount       = $currentProcessCount
            AttributedPowerW   = $sampleAttributedPower
            ResidualPowerW     = $sampleResidualPower
            AccumulatedEnergyJ = $accumulatedEnergy
            TemperatureC       = $gpuTemperature
            FanPercent         = $gpuFanUtil
            PerProcess         = $samplePerProcess
        }

        [void]$this.Samples.Add($sampleObj)

        # enforce in-memory cap
        try {
            if ($this.Samples.Count -gt $this.MaxSamplesInMemory) {
                $removeCount = $this.Samples.Count - $this.MaxSamplesInMemory
                # RemoveRange exists on List<T>
                $this.Samples.RemoveRange(0, $removeCount)
            }
        } catch {}

        # ---------- realtime CSV logging (cached writers & batched flush) ----------
        try {
            # Writers were created in constructor where possible; Sample() will fall back to lazy open if necessary
            if ($null -eq $this.SamplesCsvWriter -or $null -eq $this.SamplesCsvWriter.BaseStream) {
                # lazy-open (same logic as constructor) - rarely runs since constructor tries to open
                $timestampForFile = $this.TimeStampLogging.ToString('yyyyMMdd_HHmmss')
                $outSpec = $this.DiagnosticsOutputPath
                if (-not $outSpec) {
                    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
                    $outSpec = Join-Path $scriptDir 'power_diagnostics'
                }

                if ($outSpec -match '\.csv$') {
                    $csvPath = $outSpec
                    $csvPathProcesses = $outSpec -replace '\.csv$','_processes.csv'
                    $diagDir = Split-Path -Parent $csvPath
                    if ($diagDir -and -not (Test-Path $diagDir)) { New-Item -ItemType Directory -Path $diagDir -Force | Out-Null }
                } else {
                    $diagDir = $outSpec
                    if (-not (Test-Path $diagDir)) { New-Item -ItemType Directory -Path $diagDir -Force | Out-Null }
                    $csvPath = Join-Path $diagDir ("samples_$timestampForFile.csv")
                    $csvPathProcesses = Join-Path $diagDir ("processes_$timestampForFile.csv")
                }

                $needHeader = -not (Test-Path $csvPath)
                $fsSamples = [System.IO.FileStream]::new($csvPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
                $fsSamples.Seek(0, [System.IO.SeekOrigin]::End) | Out-Null
                $this.SamplesCsvWriter = [System.IO.StreamWriter]::new($fsSamples, [System.Text.Encoding]::UTF8)
                $this.SamplesCsvWriter.AutoFlush = $false
                if ($needHeader -and ($fsSamples.Length -eq 0)) {
                    $this.SamplesCsvWriter.WriteLine('Timestamp,PowerW,ActivePowerW,ExcessPowerW,GpuSMUtil,GpuMemUtil,GpuEncUtil,GpuDecUtil,GpuWeightTotal,ProcessWeightTotal,ProcessCount,AttributedPowerW,ResidualPowerW,AccumulatedEnergyJ,TemperatureC,FanPercent')
                }
            }

            # write sample line (single concatenated string) — round here only
            $line = '"' + $tsEsc_local + '",' +
                    ([double]$gpuPower).ToString('F4', $invariantCulture) + ',' +
                    ([double]$gpuActivePower).ToString('F4', $invariantCulture) + ',' +
                    ([double]$gpuExcessPower).ToString('F4', $invariantCulture) + ',' +
                    ([double]$gpuSmUtil).ToString('F2', $invariantCulture) + ',' +
                    ([double]$gpuMemUtil).ToString('F2', $invariantCulture) + ',' +
                    ([double]$gpuEncUtil).ToString('F2', $invariantCulture) + ',' +
                    ([double]$gpuDecUtil).ToString('F2', $invariantCulture) + ',' +
                    ([double]$gpuWeightedTotal).ToString('F2', $invariantCulture) + ',' +
                    ([double]$processWeightTotal).ToString('F2', $invariantCulture) + ',' +
                    [int]$currentProcessCount + ',' +
                    ([double]$sampleAttributedPower).ToString('F6', $invariantCulture) + ',' +
                    ([double]$sampleResidualPower).ToString('F6', $invariantCulture) + ',' +
                    ([double]$accumulatedEnergy).ToString('F8', $invariantCulture) + ',' +
                    ([double]$gpuTemperature).ToString('F1', $invariantCulture) + ',' +
                    ([double]$gpuFanUtil).ToString('F1', $invariantCulture)

            $this.SamplesCsvWriter.WriteLine($line)

            # open and cache per-process writer if missing
            if ($null -eq $this.ProcessesCsvWriter -or $null -eq $this.ProcessesCsvWriter.BaseStream) {
                $timestampForFile = $this.TimeStampLogging.ToString('yyyyMMdd_HHmmss')
                $outSpec = $this.DiagnosticsOutputPath
                if (-not $outSpec) {
                    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
                    $outSpec = Join-Path $scriptDir 'power_diagnostics'
                }

                if ($outSpec -match '\.csv$') {
                    $csvPathProcesses = $outSpec -replace '\.csv$','_processes.csv'
                } else {
                    $diagDir = $outSpec
                    $csvPathProcesses = Join-Path $diagDir ("processes_$timestampForFile.csv")
                }

                $needHeaderProc = -not (Test-Path $csvPathProcesses)
                $fsProcs = [System.IO.FileStream]::new($csvPathProcesses, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
                $fsProcs.Seek(0, [System.IO.SeekOrigin]::End) | Out-Null
                $this.ProcessesCsvWriter = [System.IO.StreamWriter]::new($fsProcs, [System.Text.Encoding]::UTF8)
                $this.ProcessesCsvWriter.AutoFlush = $false
                if ($needHeaderProc -and ($fsProcs.Length -eq 0)) {
                    $this.ProcessesCsvWriter.WriteLine('Timestamp,PID,ProcessName,SMUtil,MemUtil,EncUtil,DecUtil,PowerW,EnergyJ,AccumulatedEnergyJ,WeightedUtil,IsIdle')
                }
            }

            # write per-process lines (round here)
            foreach ($perProcessEntry in $samplePerProcess.Values) {
                $pnameEsc = ($perProcessEntry.ProcessName -replace '"','""') -replace '\r|\n',' '
                $line2 = '"' + $tsEsc_local + '",' +
                        $perProcessEntry.PID + ',"' + $pnameEsc + '",' +
                        ([double]$perProcessEntry.SMUtil).ToString('F2', $invariantCulture) + ',' +
                        ([double]$perProcessEntry.MemUtil).ToString('F2', $invariantCulture) + ',' +
                        ([double]$perProcessEntry.EncUtil).ToString('F2', $invariantCulture) + ',' +
                        ([double]$perProcessEntry.DecUtil).ToString('F2', $invariantCulture) + ',' +
                        ([double]$perProcessEntry.PowerW).ToString('F6', $invariantCulture) + ',' +
                        ([double]$perProcessEntry.EnergyJ).ToString('F8', $invariantCulture) + ',' +
                        ([double]$perProcessEntry.AccumulatedProcessEnergyJ).ToString('F8', $invariantCulture) + ',' +
                        ([double]$perProcessEntry.WeightedUtil).ToString('F4', $invariantCulture) + ',' +
                        ([bool]$perProcessEntry.IsIdle).ToString()
                $this.ProcessesCsvWriter.WriteLine($line2)
            }

            # batched flush (reduce I/O churn)
            $this.SamplesSinceLastFlush += 1
            if ($this.SamplesSinceLastFlush -ge $this.SamplesFlushInterval) {
                try {
                    $this.SamplesCsvWriter.Flush()
                    $this.ProcessesCsvWriter.Flush()
                } catch {}
                $this.SamplesSinceLastFlush = 0
            }
        }
        catch {
            # do not break sampling
            Write-Log("Realtime CSV logging failed: $($_.Exception.Message)")
        }
    }

    # Run method with fixed-rate scheduling and safe writer close on exit
    [void] Run([Nullable[int]]$Duration) {
        # Prompt user to start workload
        Write-Host ""; Write-Log("READY: Start the workload now.")
        Write-Host "Press ENTER when the workload is running and you want to begin sampling..."
        [void][System.Console]::ReadLine()

        # Ensure high-resolution stopwatch is available and warm
        if (-not $this.Stopwatch) {
            $this.Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        } else {
            $this.Stopwatch.Restart()
        }

        if (-not $this.StopwatchFrequency) {
            $this.StopwatchFrequency = [int64][System.Diagnostics.Stopwatch]::Frequency
        }

        # compute interval in ticks
        $intervalTicks = [int64]([double]$this.SampleIntervalMs * $this.StopwatchFrequency / 1000.0)
        if ($intervalTicks -le 0) { $intervalTicks = 1 }

        # set initial scheduled tick (first sample scheduled after interval)
        $this.LastSampleTime = $this.Stopwatch.ElapsedTicks
        $scheduledTick = $this.LastSampleTime + $intervalTicks

        # duration ticks if provided
        $runForever = $false
        $endTicks = 0
        if ($null -eq $Duration) {
            $runForever = $true
            Write-Log("Monitoring GPU indefinitely (no duration specified). Use Ctrl+C to stop.")
        } else {
            $endTicks = [int64]($Duration * $this.StopwatchFrequency)
            Write-Log("Monitoring GPU for $Duration seconds...")
            $endWallClock = (Get-Date).AddSeconds([double]$Duration)
            Write-Log("End time (wall clock): $endWallClock")
            # Convert end ticks to absolute stopwatch ticks
            $endTicks = $this.Stopwatch.ElapsedTicks + $endTicks
        }

        $sampleCount = 0
        try {
            while ($runForever -or ($this.Stopwatch.ElapsedTicks -lt $endTicks)) {
                $now = $this.Stopwatch.ElapsedTicks
                $remaining = $scheduledTick - $now
                if ($remaining -le 0) {
                    # it's time (or past): do sample
                    $this.Sample()
                    $sampleCount++

                    # advance scheduledTick forward to the next future schedule
                    do {
                        $scheduledTick += $intervalTicks
                        $now = $this.Stopwatch.ElapsedTicks
                    } while ($scheduledTick -le $now)
                } else {
                    # sleep most of the remaining time to save CPU; wake a little earlier for accuracy
                    $remainingMs = [int]([double]$remaining * 1000.0 / $this.StopwatchFrequency)
                    if ($remainingMs -gt 5) {
                        Start-Sleep -Milliseconds ($remainingMs - 3)
                    } else {
                        # very short remaining time — yield to OS
                        [System.Threading.Thread]::Sleep(0)
                    }
                }
            }
        }
        catch [System.Exception] {
            Write-Log("Monitoring stopped: $($_.Exception.Message)")
        }
        finally {
            # flush and close writers to ensure files are written (Ctrl+C and shutdown)
            try { $this.CloseWriters() } catch {}

            $Duration = $this.Stopwatch.Elapsed.TotalSeconds
            Write-Log("Collected $sampleCount samples over $Duration seconds.")
            try { $this.Report($Duration) } catch {}
        }
    }

    [void] Report([double]$Duration) {
        Write-Host "`n==== GPU POWER SUMMARY ===="

        $safeDuration = if ($Duration -le 0 -or [double]::IsNaN($Duration)) { 1.0 } else { [double]$Duration }
        $avgGpuPower = if ($safeDuration -gt 0) { $this.GpuEnergyJoules / $safeDuration } else { 0.0 }
        Write-Host ("Total GPU - Accumulative: {0,8:F2} J  |  Average: {1,6:F2} W" -f $this.GpuEnergyJoules, $avgGpuPower)

        Write-Host "`n==== PROCESS POWER ATTRIBUTION (Multi-Metric) ===="
        if ($this.ProcessEnergyJoules.Count -eq 0) {
            Write-Host "No process energy data recorded."
            return
        }

        # Build list of entries (Key,Value) to avoid repeated enumerations
        $entriesList = New-Object 'System.Collections.Generic.List[object]'
        $iter = $this.ProcessEnergyJoules.GetEnumerator()
        while ($iter.MoveNext()) {
            $entriesList.Add(@{ Key = $iter.Current.Key; Value = [double]$iter.Current.Value })
        }

        # Map PIDs -> process names with a single Get-Process call (best-effort)
        $allPids = $entriesList | ForEach-Object { [int]$_.Key } | Sort-Object -Unique
        $pidToName = @{}
        if ($allPids.Count -gt 0) {
            try {
                $procs = Get-Process -Id $allPids -ErrorAction SilentlyContinue
                foreach ($processObj in $procs) { $pidToName[[int]$processObj.Id] = $processObj.ProcessName }
            } catch {}
        }

        # Fallback: if Get-Process didn't return a name (process ended), try to obtain the name
        if ($allPids.Count -gt 0) {
            foreach ($procId in $allPids) {
                if (-not $pidToName.ContainsKey($procId)) {
                    $foundName = $null
                    foreach ($sampleItem in $this.Samples) {
                        if ($sampleItem.PerProcess -and $sampleItem.PerProcess.ContainsKey([string]$procId)) {
                            $perProcessEntry = $sampleItem.PerProcess[[string]$procId]
                            if ($perProcessEntry -and $perProcessEntry.ProcessName) {
                                $foundName = [string]$perProcessEntry.ProcessName
                                break
                            }
                        }
                    }
                    if ($foundName) { $pidToName[[int]$procId] = $foundName }
                }
            }
        }

        # Sum per-process energies
        $sumTotalProcessesEnergy = 0.0
        foreach ($entryPair in $entriesList) { $sumTotalProcessesEnergy += [double]$entryPair.Value }

        # Diagnostic compare total vs sum of processes
        if ($sumTotalProcessesEnergy -gt 0) {
            $diff = [math]::Abs([double]$this.GpuEnergyJoules - $sumTotalProcessesEnergy)
            $diffPct = if ($this.GpuEnergyJoules -gt 0) { 100.0 * $diff / $this.GpuEnergyJoules } else { 0.0 }
            Write-Log("DIAGNOSTIC: Measured total {0:F2} J vs summed attributed {1:F2} J -> Diff: {2:F2} J ({3:F1}%)" -f $this.GpuEnergyJoules, $sumTotalProcessesEnergy, $diff, $diffPct)
        }

        # Aggregate by process name and print
        $aggregatedByProcess = @{}
        foreach ($entryPair in $entriesList) {
            $pidNum = [int]$entryPair.Key
            $pname = if ($pidToName.ContainsKey($pidNum)) { $pidToName[$pidNum] } else { "[exited] PID $pidNum" }
            if (-not $aggregatedByProcess.ContainsKey($pname)) { $aggregatedByProcess[$pname] = @{ Energy = 0.0; Pids = New-Object 'System.Collections.Generic.List[int]' } }
            $aggregatedByProcess[$pname].Energy += [double]$entryPair.Value
            [void]$aggregatedByProcess[$pname].Pids.Add($pidNum)
        }

        $aggregatedByProcess.GetEnumerator() | Sort-Object { $_.Value.Energy } -Descending | ForEach-Object {
            $name = $_.Key
            $energy = $_.Value.Energy
            $avgPowerProc = if ($safeDuration -gt 0) { $energy / $safeDuration } else { 0.0 }
            $percentageOfTotalEnergy = if ($sumTotalProcessesEnergy -gt 0) { 100.0 * $energy / $sumTotalProcessesEnergy } else { 0.0 }
            $pidList = ($_.Value.Pids -join ',')
            Write-Host ("{0,-30} Accumulative: {1,8:F2} J  |  Average: {2,6:F2} W  ({3,5:F1}%)  PIDs: {4}" -f $name, $energy, $avgPowerProc, $percentageOfTotalEnergy, $pidList)
        }
    }
}

# Main execution
try {
    $monitor = [GPUProcessMonitor]::new($SampleInterval, $WeightSM, $WeightMem, $WeightEnc, $WeightDec, $DiagnosticsOutput)
    $monitor.Run($Duration)
}
catch {
    Write-Error "Error: $_"
    # ensure writers closed before exit
    try { if ($monitor) { $monitor.CloseWriters() } } catch {}
    exit 1
}