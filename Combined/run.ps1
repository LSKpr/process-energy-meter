<#
.SYNOPSIS
    Interactive Process Power Monitor
.DESCRIPTION
    Command-line tool to monitor process power consumption with:
    - Top X processes view
    - Detailed single process monitoring with power graph
    - Real-time power consumption tracking
    
    Requires: Administrator privileges and LibreHardwareMonitorLib.dll

    Flags of interest:
    -MeasurementIntervalSeconds: how often to sample CPU in seconds (default 2)
    -SampleInterval: how often to sample GPU in ms (default 100)
    -WeightSM, WeightMem, WeightEnc, WeightDec: weights for the attribution formula (defaults: 1.0, 0.5, 0.25, 0.15)
    -DiagnosticsOutput: base path for writing diagnostics CSV files (default "power_diagnostics")

.EXAMPLE
    .\run.ps1
    Starts the interactive monitor with default settings.

    .\run.ps1 -MeasurementIntervalSeconds 1 -SampleInterval 50 -WeightSM 1.0 -WeightMem 0.5 -WeightEnc 0.25 -WeightDec 0.15
#>

param(
    [int]$MeasurementIntervalSeconds = 2,
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

# Dot-source logic
. .\CpuMonitor.ps1
. .\GpuMonitor.ps1

#----------------------------------------------
#region Utility
#----------------------------------------------

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

#----------------------------------------------
#region Interface
#----------------------------------------------

$script:CurrentTopList = @()
$script:CurrentView = "list"
$script:CurrentViewParam = 20
$script:MonitoringStartTime = Get-Date
$script:MeasurementInterval = 2

$script:GpuMonitor = $null

function Show-TopList {
    param(
        [int]$Count,
        [bool]$ClearScreen = $true,
        [int]$SampleCount
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

    Write-Host "`n==== GPU POWER SUMMARY ====" -ForegroundColor $script:Colors.Header
    try { $script:GpuMonitor.Report($SampleCount) } catch {}
    
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
    Write-Host "GPU: $($script:GpuMonitor.GpuName)" -ForegroundColor $script:Colors.Value
    
    $testSystemPower = Get-SystemPowerConsumption
    if ($null -eq $testSystemPower) {
        Write-Host "System Power: Not Available" -ForegroundColor Yellow
    }
    else {
        Write-Host "System Power: Available" -ForegroundColor Green
    }
    
    Write-Host ""
    # Measure idle metrics
    Write-Host "Measuring idle GPU power..."
    $script:GpuMonitor.MeasureIdleMetrics()
    Write-Host "Collecting initial data..." -ForegroundColor $script:Colors.Info
    
    # Collect initial data
    for ($i = 0; $i -lt 3; $i++) {
        $null = Update-ProcessEnergyData -IntervalSeconds $IntervalSeconds
        Start-Sleep -Seconds $IntervalSeconds
        Write-Host "." -NoNewline -ForegroundColor $script:Colors.Info
    }

    Write-Host "`n`nStarting with list 20 view...`n" -ForegroundColor Green
    Start-Sleep -Seconds 1
    
    # Ensure high-resolution stopwatch is available and warm
    if (-not $script:GpuMonitor.Stopwatch) {
        $script:GpuMonitor.Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    } else {
        $script:GpuMonitor.Stopwatch.Restart()
    }

    if (-not $script:GpuMonitor.StopwatchFrequency) {
        $script:GpuMonitor.StopwatchFrequency = [int64][System.Diagnostics.Stopwatch]::Frequency
    }

    # compute interval in ticks
    $intervalTicks = [int64]([double]$script:GpuMonitor.SampleIntervalMs * $script:GpuMonitor.StopwatchFrequency / 1000.0)
    if ($intervalTicks -le 0) { $intervalTicks = 1 }

    # set initial scheduled tick (first sample scheduled after interval)
    $script:GpuMonitor.LastSampleTime = $script:GpuMonitor.Stopwatch.ElapsedTicks
    $scheduledTick = $script:GpuMonitor.LastSampleTime + $intervalTicks

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

    $sampleCount = 0
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
                            Show-TopList -Count $script:CurrentViewParam -ClearScreen $true -
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
            
            $now = $script:GpuMonitor.Stopwatch.ElapsedTicks
            $remaining = $scheduledTick - $now
            if ($remaining -le 0) {
                # it's time (or past): do sample
                $script:GpuMonitor.Sample()
                $sampleCount++

                # advance scheduledTick forward to the next future schedule
                do {
                    $scheduledTick += $intervalTicks
                    $now = $script:GpuMonitor.Stopwatch.ElapsedTicks
                } while ($scheduledTick -le $now)
            } else {
                # sleep most of the remaining time to save CPU; wake a little earlier for accuracy
                $remainingMs = [int]([double]$remaining * 1000.0 / $script:GpuMonitor.StopwatchFrequency)
                if ($remainingMs -gt 5) {
                    Start-Sleep -Milliseconds ($remainingMs - 3)
                } else {
                    # very short remaining time — yield to OS
                    [System.Threading.Thread]::Sleep(0)
                }
            }
        }
    }
    finally {
        # Cleanup
        [Console]::CursorVisible = $true
        
        if ($_.Exception.Message -eq "EXIT") {
            Write-Host "`nExiting..." -ForegroundColor $script:Colors.Info
        }

        # flush and close writers to ensure files are written (Ctrl+C and shutdown)
        try { $script:GpuMonitor.CloseWriters() } catch {}
    }
}

#endregion

#----------------------------------------------
#region Main Execution
#----------------------------------------------

# Check admin privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "`nERROR: This script requires Administrator privileges!" -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again.`n" -ForegroundColor Yellow
    exit 1
}

Clear-Host
Write-Host "`nInitializing Interactive Process Power Monitor..." -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Cyan

$initialized = Initialize-LibreHardwareMonitor

if (-not $initialized) {
    Write-Host "`nFailed to initialize. Exiting..." -ForegroundColor Red
    exit 1
}

try {
    $script:GpuMonitor = [GPUProcessMonitor]::new($SampleInterval, $WeightSM, $WeightMem, $WeightEnc, $WeightDec, $DiagnosticsOutput)

    Write-Host "[OK] Hardware monitoring ready" -ForegroundColor Green
    
    Start-CommandLineMode -IntervalSeconds $MeasurementIntervalSeconds
}
catch {
    Write-Error "Error: $_"
    # ensure writers closed before exit
    try { if ($script:GpuMonitor) { $script:GpuMonitor.CloseWriters() } } catch {}
}

# Cleanup
if ($null -ne $script:Computer) {
    $script:Computer.Close()
}

Write-Host "`nThank you for using Process Power Monitor!`n" -ForegroundColor Cyan

#endregion