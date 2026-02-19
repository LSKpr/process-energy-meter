param(
    [int]$MeasurementIntervalSeconds = 2,
    [int]$SampleInterval = 100,
    [double]$WeightSMP = 1.0,
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

function Read-LatestCpuProcessAggregates {
    param(
        [string]$CpuProcessesCsvPath,
        [int]$TailLines = 1024
    )

    $result = @{}

    if (-not $CpuProcessesCsvPath) { return $result }
    if (-not (Test-Path $CpuProcessesCsvPath)) { return $result }

    # read last N lines (Get-Content -Tail is efficient)
    $lines = Get-Content -Path $CpuProcessesCsvPath -Tail $TailLines -ErrorAction SilentlyContinue
    if (-not $lines) { return $result }

    # Regex that accepts quoted or unquoted timestamp, quoted or unquoted process name,
    # then four numeric fields. Captures name, cpu, power, energy, accumulated.
    $pattern = '^\s*(?:"(?<ts>[^"]*)"|(?<ts>[^,]*))\s*,\s*(?:"(?<name>(?:[^"]|"")*)"|(?<name>[^,]*))\s*,\s*(?<cpu>[^,]+)\s*,\s*(?<pw>[^,]+)\s*,\s*(?<ej>[^,]+)\s*,\s*(?<acc>[^,]+)\s*$'
    $regex = [System.Text.RegularExpressions.Regex]::new($pattern, [System.Text.RegularExpressions.RegexOptions]::Compiled)

    $ci = [System.Globalization.CultureInfo]::InvariantCulture

    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        # skip header lines that begin with "Timestamp" (common header)
        if ($line.TrimStart().StartsWith('Timestamp', [System.StringComparison]::OrdinalIgnoreCase)) { continue }

        $m = $regex.Match($line)
        if (-not $m.Success) {
            # skip lines that don't match expected schema
            continue
        }

        $procNameRaw = $m.Groups['name'].Value
        # unescape double quotes inside CSV quoted names
        $procName = $procNameRaw -replace '""', '"'

        # parse numeric fields with invariant culture, defensively
        $cpuPct = 0.0
        $pw = 0.0
        $ej = 0.0
        $acc = 0.0
        try { $cpuPct = [double]::Parse($m.Groups['cpu'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $pw = [double]::Parse($m.Groups['pw'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $ej = [double]::Parse($m.Groups['ej'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $acc = [double]::Parse($m.Groups['acc'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}

        # last occurrence wins (tail is chronological), so later lines overwrite earlier ones
        $result[$procName] = @{
            LastSeenCpu = $cpuPct
            LastSeenPowerMw = $pw
            CpuEnergy = $acc
        }
    }

    return $result
}


function Read-LatestCpuSample {
    param(
        [string]$CpuSamplesCsvPath,
        [int]$TailLines = 10
    )

    if (-not $CpuSamplesCsvPath) { return $null }
    if (-not (Test-Path $CpuSamplesCsvPath)) { return $null }

    $lines = Get-Content -Path $CpuSamplesCsvPath -Tail $TailLines -ErrorAction SilentlyContinue
    if (-not $lines) { return $null }

    # Regex for a CPU sample row:
    # Timestamp,CpuPowerMw,SystemPowerMw,TotalCpuPercent,MeasurementIntervalSeconds,AccumulatedCpuEnergymJ
    $pattern = '^\s*(?:"(?<ts>[^"]*)"|(?<ts>[^,]*))\s*,\s*(?<cpuMw>[^,]+)\s*,\s*(?<sysMw>[^,]+)\s*,\s*(?<totalCpu>[^,]+)\s*,\s*(?<interval>[^,]+)\s*,\s*(?<acc>[^,]+)\s*$'
    $regex = [System.Text.RegularExpressions.Regex]::new($pattern, [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ci = [System.Globalization.CultureInfo]::InvariantCulture

    # find last matching non-header line (iterate backwards)
    for ($i = $lines.Length - 1; $i -ge 0; $i--) {
        $line = $lines[$i]
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line.TrimStart().StartsWith('Timestamp', [System.StringComparison]::OrdinalIgnoreCase)) { continue }

        $m = $regex.Match($line)
        if (-not $m.Success) { continue }

        $ts = $m.Groups['ts'].Value
        $cpuMw = 0.0; $sysMw = 0.0; $totalCpu = 0.0; $interval = 0.0; $acc = 0.0
        try { $cpuMw = [double]::Parse($m.Groups['cpuMw'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $sysMw = [double]::Parse($m.Groups['sysMw'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $totalCpu = [double]::Parse($m.Groups['totalCpu'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $interval = [double]::Parse($m.Groups['interval'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $acc = [double]::Parse($m.Groups['acc'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}

        return @{
            Timestamp = $ts
            CpuPowerMw = $cpuMw
            SystemPowerMw = $sysMw
            TotalCpuPercent = $totalCpu
            IntervalSeconds = $interval
            AccumulatedCpuEnergymJ = $acc
        }
    }

    return $null
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

#region Command Line Interface

function Show-TopList {
    param(
        [int]$Count,
        [bool]$ClearScreen = $true,
        [int]$SampleCount
    )
    
    if ($Count -lt 1) { $Count = 1 }
    if ($Count -gt 50) { $Count = 50 }
    
    $elapsed = if ($script:Stopwatch) { $script:Stopwatch.Elapsed.ToString("hh\:mm\:ss") } else { (Get-Date -UFormat "%H:%M:%S") }
    
    # Determine CPU CSV paths (fall back to existing script globals if present)
    $cpuProcsPath = if ($script:CpuProcessesCsvPath) { $script:CpuProcessesCsvPath } else { $null }
    $cpuSamplesPath = if ($script:CpuSamplesCsvPath) { $script:CpuSamplesCsvPath } else { $null }

    # Try to read per-process aggregates from CSV; if not available, keep existing in-memory map
    $processAggregates = @{}
    if ($cpuProcsPath -and (Test-Path $cpuProcsPath)) {
        try {
            $processAggregates = Read-LatestCpuProcessAggregates -CpuProcessesCsvPath $cpuProcsPath -TailLines 200
        } catch {
            $processAggregates = @{}
        }
    }

    # build $script:ProcessEnergyData-like map for UI use:
    if ($processAggregates.Count -gt 0) {
        $script:ProcessEnergyData = @{}
        foreach ($kv in $processAggregates.GetEnumerator()) {
            $pname = $kv.Key
            $pdata = $kv.Value
            $script:ProcessEnergyData[$pname] = @{
                CpuEnergy = $pdata.CpuEnergy
                SystemEnergy = 0
                LastSeenCpu = $pdata.LastSeenCpu
                LastSeenPowerMw = $pdata.LastSeenPowerMw
                PowerHistory = New-Object System.Collections.ArrayList
            }
        }
    } else {
        # if CSV not available, ensure ProcessEnergyData exists (use existing in-memory if present)
        if (-not $script:ProcessEnergyData) { $script:ProcessEnergyData = @{} }
    }

    # overall sample (CPU)
    if ($cpuSamplesPath -and (Test-Path $cpuSamplesPath)) {
        try {
            $latestSample = Read-LatestCpuSample -CpuSamplesCsvPath $cpuSamplesPath -TailLines 10
            if ($latestSample) {
                $script:CurrentCpuPower = $latestSample.CpuPowerMw
                $script:CurrentSystemPower = $latestSample.SystemPowerMw
                $script:CurrentTotalCpuPercent = $latestSample.TotalCpuPercent
            }
        } catch {
            # ignore parsing failure and keep previous values
        }
    }

    # Calculate total energy (based on the data we have in $script:ProcessEnergyData)
    $totalCpuEnergy = 0.0
    $totalSystemEnergy = 0.0
    foreach ($value in $script:ProcessEnergyData.Values) {
        $totalCpuEnergy += [double]$value.CpuEnergy
        $totalSystemEnergy += [double]$value.SystemEnergy
    }

    if ($ClearScreen) {
        Clear-Host
    } else {
        # Move to top-left without clearing
        [Console]::SetCursorPosition(0, 0)
    }
    
    Write-Host "`n========================================================" -ForegroundColor $script:Colors.Header
    Write-Host "   CPU Process Power Monitor - Top $Count                " -ForegroundColor $script:Colors.Header
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

    #GPU section (unchanged)
    # Build list of entries (Key,Value) to avoid repeated enumerations
    $entriesList = New-Object 'System.Collections.Generic.List[object]'
    $iter = $script:ProcessEnergyJoules.GetEnumerator()
    while ($iter.MoveNext()) {
        $entriesList.Add(@{ Key = $iter.Current.Key; Value = [double]$iter.Current.Value })
    }

    # Map PIDs -> process names with a single Get-Process call (best-effort)
    $allPids = $entriesList | ForEach-Object { [int]$_.Key } | Sort-Object -Unique
    $pidToName = @{}
    if ($allPids.Count -gt 0) {
        try {
            $procs = Get-Process -Id $allPids -ErrorAction SilentlyContinue
            foreach ($pr in $procs) { $pidToName[[int]$pr.Id] = $pr.ProcessName }
        } catch {}
    }

    # Efficient fallback: build a single fast map from samples for missing PIDs
    # (avoid repeated scanning of $script:Samples per PID)
    $missingPids = @()
    foreach ($procId in $allPids) { if (-not $pidToName.ContainsKey($procId)) { $missingPids += $procId } }
    if ($missingPids.Count -gt 0) {
        $samplePidNameMap = @{}
        foreach ($s in $script:Samples) {
            if ($s.PerProcess) {
                foreach ($pidKey in $s.PerProcess.Keys) {
                    $pidInt = [int]$pidKey
                    if (-not $pidToName.ContainsKey($pidInt) -and -not $samplePidNameMap.ContainsKey($pidInt)) {
                        $pp = $s.PerProcess[[string]$pidKey]
                        if ($pp -and $pp.ProcessName) {
                            $samplePidNameMap[$pidInt] = [string]$pp.ProcessName
                        }
                    }
                }
            }
            # short-circuit if we've resolved all missing pids
            if ($samplePidNameMap.Keys.Count -ge $missingPids.Count) { break }
        }
        foreach ($mp in $samplePidNameMap.Keys) { if (-not $pidToName.ContainsKey($mp)) { $pidToName[$mp] = $samplePidNameMap[$mp] } }
    }

    Write-Host "`n========================================================" -ForegroundColor $script:Colors.Header
    Write-Host "   GPU Process Power Monitor                    " -ForegroundColor $script:Colors.Header
    Write-Host "========================================================" -ForegroundColor $script:Colors.Header
    Write-Host ""
    Write-Host ("  Measurements:        {0}" -f $SampleCount) -ForegroundColor $script:Colors.Value
    Write-Host ("  Measurement Interval: {0}ms" -f $script:SampleIntervalMs) -ForegroundColor $script:Colors.Value
    Write-Host ("  Tracked Processes:   {0}" -f $allPids.Count) -ForegroundColor $script:Colors.Value
    Write-Host ("  GPU Power attribution weights: SM={0:F2}, Mem={1:F2}, Enc={2:F2}, Dec={3:F2}" -f $script:WeightSM, $script:WeightMemory, $script:WeightEncoder, $script:WeightDecoder)
    Write-Host ("  Idle GPU power measured: {0:N2}W [Min: {1:N2}W  Max:{2:N2}W] with {3} Processes; Temp: {4:F1}C  Fan: {5:F1}%" -f `
        $script:GpuIdlePower, $script:GpuIdlePowerMin, $script:GpuIdlePowerMax, $script:IdleProcessCount, $script:GpuIdleTemperatureC, $script:GpuIdleFanPercent)
    Write-Host ""
    Write-Host ("  Total GPU Energy:    {0:F2} J" -f $script:GpuEnergyJoules) -ForegroundColor Yellow

    # Sum per-process energies
    $sumTotalProcessesEnergy = 0.0
    foreach ($e in $entriesList) { $sumTotalProcessesEnergy += [double]$e.Value }

    # Diagnostic compare total vs sum of processes
    if ($sumTotalProcessesEnergy -gt 0) {
        $diff = [math]::Abs([double]$script:GpuEnergyJoules - $sumTotalProcessesEnergy)
        $diffPct = if ($script:GpuEnergyJoules -gt 0) { 100.0 * $diff / $script:GpuEnergyJoules } else { 0.0 }
        Write-Host ("  Measured total {0:F2} J vs summed attributed {1:F2} J -> Diff: {2:F2} J ({3:F1}%)" -f $script:GpuEnergyJoules, $sumTotalProcessesEnergy, $diff, $diffPct)
    }
    else {
        Write-Host ""
    }
    Write-Host ""
    Write-Host "Top GPU Processes:" -ForegroundColor $script:Colors.Highlight
    Write-Host ("{0,-15} {1,-30} {2,12} {3,10}" -f "PID", "Process", "GPU Energy", "GPU %") -ForegroundColor $script:Colors.Header
    Write-Host ("-" * 75) -ForegroundColor $script:Colors.Header

    # Build per-PID list and sort descending by energy (do not aggregate by name)
    $perPidList = @()
    foreach ($e in $entriesList) {
        $pidNum = [int]$e.Key
        $pname = if ($pidToName.ContainsKey($pidNum)) { $pidToName[$pidNum] } else { "[exited] PID $pidNum" }
        $perPidList += [PSCustomObject]@{
            PID = $pidNum
            ProcessName = $pname
            Energy = [double]$e.Value
        }
    }

    $sortedPerPid = $perPidList | Sort-Object -Property @{ Expression = { $_.Energy } } -Descending

    foreach ($entry in $sortedPerPid) {
        $pidText = $entry.PID
        $name = $entry.ProcessName
        $energy = $entry.Energy
        $pct = if ($sumTotalProcessesEnergy -gt 0) { 100.0 * $energy / $sumTotalProcessesEnergy } else { 0.0 }
        Write-Host ("{0,-15} {1,-30} {2,12:F2}J {3,5:F1}%" -f $pidText, $name, $energy, $pct)
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
    Write-Host "GPU: $($script:GpuName)" -ForegroundColor $script:Colors.Value
    Write-Host "========================================================" -ForegroundColor $script:Colors.Header
    
    $testSystemPower = Get-SystemPowerConsumption
    if ($null -eq $testSystemPower) {
        Write-Host "System Power: Not Available" -ForegroundColor Yellow
    }
    else {
        Write-Host "System Power: Available" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "Collecting initial CPU data..." -ForegroundColor $script:Colors.Info

    # Ensure high-resolution stopwatch is available and warm before CPU warm-up
    if (-not $script:Stopwatch) {
        $script:Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    } else {
        $script:Stopwatch.Restart()
    }

    if (-not $script:StopwatchFrequency) {
        $script:StopwatchFrequency = [int64][System.Diagnostics.Stopwatch]::Frequency
    }

    # reset CPU sampling baseline so Update-CPUEnergyData will initialize on first call
    $script:LastCpuSampleTicks = 0

    # Collect initial CPU data (warm-up); Update-CPUEnergyData uses stopwatch dt internally
    for ($i = 0; $i -lt 3; $i++) {
        $null = Update-CPUEnergyData
        Start-Sleep -Seconds $IntervalSeconds
        Write-Host "." -NoNewline -ForegroundColor $script:Colors.Info
    }

    # Measure idle GPU metrics
    Write-Host "`nMeasuring idle GPU power (please ensure GPU is idle; close heavy apps)."
    Measure-IdleMetrics
    
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

    # compute GPU sample interval in ticks (SampleIntervalMs is milliseconds)
    $intervalTicksGPU = [int64]([double]$script:SampleIntervalMs * $script:StopwatchFrequency / 1000.0)
    if ($intervalTicksGPU -le 0) { $intervalTicksGPU = 1 }

    # compute CPU measurement interval in ticks (MeasurementInterval is seconds)
    $intervalTicksCPU = [int64]([double]$script:MeasurementInterval * $script:StopwatchFrequency)
    if ($intervalTicksCPU -le 0) { $intervalTicksCPU = 1 }

    # Track last update times (ticks)
    $script:LastSampleTime = $script:Stopwatch.ElapsedTicks

    $scheduledTickGPU = $script:LastSampleTime + $intervalTicksGPU
    $scheduledTickCPU = $script:LastSampleTime + $intervalTicksCPU

    # 2 seconds for display refresh -> 2000 ms converted to ticks
    $scheduledTickDisplay = $script:LastSampleTime + [int64]([double]2000 * $script:StopwatchFrequency / 1000.0)
    
    $sampleCount = 0
    try {
        # Show initial display
        Show-TopList -Count $script:CurrentViewParam -ClearScreen $true
        [Console]::CursorVisible = $true
        
        # Main input loop
        while ($true) {
            $now = $script:Stopwatch.ElapsedTicks

            $remainingGPU = $scheduledTickGPU - $now
            $remainingCPU = $scheduledTickCPU - $now
            $remainingDisplay = $scheduledTickDisplay - $now

            # GPU sampling
            if ($remainingGPU -le 0) {
                Update-GPUEnergyData
                $sampleCount++

                # advance scheduledTickGPU forward to the next future schedule
                do {
                    $scheduledTickGPU += $intervalTicksGPU
                    $now = $script:Stopwatch.ElapsedTicks
                } while ($scheduledTickGPU -le $now)
            }

            # CPU sampling (stopwatch-driven sampler Update-CPUEnergyData)
            if ($remainingCPU -le 0) {
                Update-CPUEnergyData

                # advance scheduledTickCPU forward to the next future schedule
                do {
                    $scheduledTickCPU += $intervalTicksCPU
                    $now = $script:Stopwatch.ElapsedTicks
                } while ($scheduledTickCPU -le $now)
            }

            # Display refresh every ~2s
            if ($remainingDisplay -le 0) {
                # Save cursor position
                $cursorLeft = [Console]::CursorLeft
                $cursorTop = [Console]::CursorTop
                
                # Update display without clearing input line
                [Console]::CursorVisible = $false
                switch ($script:CurrentView) {
                    "list" { Show-TopList -Count $script:CurrentViewParam -ClearScreen $false -SampleCount $sampleCount }
                    "focus" { Show-FocusedView -ProcessName $script:CurrentViewParam }
                }
                
                # Restore cursor position for input
                [Console]::SetCursorPosition($cursorLeft, [Console]::WindowHeight - 2)
                [Console]::CursorVisible = $true
                
                # advance scheduledTickDisplay forward to the next future schedule
                do {
                    $scheduledTickDisplay += [int64]([double]2000 * $script:StopwatchFrequency / 1000.0)
                    $now = $script:Stopwatch.ElapsedTicks
                } while ($scheduledTickDisplay -le $now)
            }

            # Compute minimal remaining time (ms) to sleep to avoid busy-looping
            $remainingMsList = @()
            foreach ($rem in @(($scheduledTickGPU - $now), ($scheduledTickCPU - $now), ($scheduledTickDisplay - $now))) {
                if ($rem -gt 0) {
                    $remainingMsList += [int]([double]$rem * 1000.0 / $script:StopwatchFrequency)
                }
            }
            if ($remainingMsList.Count -gt 0) {
                $minRemainingMs = ($remainingMsList | Measure-Object -Minimum).Minimum
            } else {
                $minRemainingMs = 10
            }

            # Non-blocking input check: if key available, avoid sleeping long
            if (-not [Console]::KeyAvailable) {
                if ($minRemainingMs -gt 5) {
                    # Sleep slightly less than remaining time to wake up and be precise
                    Start-Sleep -Milliseconds ($minRemainingMs - 3)
                } else {
                    # very short remaining time — yield to OS
                    [System.Threading.Thread]::Sleep(0)
                }
            }

            # Check for user input (non-blocking)
            if ([Console]::KeyAvailable) {
                [Console]::CursorVisible = $true

                # Read the entire line using native ReadLine (this will block briefly while user types)
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
                                    # Update CPU interval ticks and reschedule next CPU sample using the new interval
                                    $intervalTicksCPU = [int64]([double]$script:MeasurementInterval * $script:StopwatchFrequency)
                                    $scheduledTickCPU = $script:Stopwatch.ElapsedTicks + $intervalTicksCPU

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
        }
    }
    finally {
        # flush and close writers to ensure files are written
        try { Close-Writers } catch {}

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

# Initialize GPU/diagnostics/writers + internal script state
Initialize-GPUProcessMonitor -SampleIntervalMs $SampleInterval -wSM $WeightSMP -wMem $WeightMem -wEnc $WeightEnc -wDec $WeightDec -diagPath $DiagnosticsOutput

# Expose CPU CSV paths for the consumer UI (keep consistent with Open-Writers)
$timestampForFile = $script:TimeStampLogging.ToString('yyyyMMdd_HHmmss')
$outSpec = $script:DiagnosticsOutputPath
if (-not $outSpec) {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $outSpec = Join-Path $scriptDir 'power_diagnostics'
}

if ($outSpec -match '\.csv$') {
    $script:CpuSamplesCsvPath = $outSpec -replace '\.csv$','_cpu_samples.csv'
    $script:CpuProcessesCsvPath = $outSpec -replace '\.csv$','_cpu_processes.csv'
} else {
    $diagDir = $outSpec
    $script:CpuSamplesCsvPath = Join-Path $diagDir ("cpu_samples_$timestampForFile.csv")
    $script:CpuProcessesCsvPath = Join-Path $diagDir ("cpu_processes_$timestampForFile.csv")
}

# Initialize LibreHardwareMonitor (CPU power)
$initialized = Initialize-LibreHardwareMonitor
if (-not $initialized) {
    Write-Host "`nFailed to initialize hardware monitoring (LibreHardwareMonitor). Exiting..." -ForegroundColor Red
    # Close any writers we opened
    try { Close-Writers } catch {}
    exit 1
}

Write-Host "[OK] Hardware monitoring ready" -ForegroundColor Green

# Start interactive mode (this will run until exit)
try {
    Start-CommandLineMode -IntervalSeconds $MeasurementIntervalSeconds
}
finally {
    # Cleanup
    try {
        if ($null -ne $script:Computer) {
            $script:Computer.Close()
        }
    } catch {}

    Write-Host "`nThank you for using Process Power Monitor!`n" -ForegroundColor Cyan
}

#endregion