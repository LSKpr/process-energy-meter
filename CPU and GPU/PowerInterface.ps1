# ===========================
# PowerInterface.ps1
# Interactive UI (CONSUMER)
# ===========================

# NOTE:
# Core script-scoped state (SampleIntervalMs, writers, caches, stopwatch, etc.)
# is owned and initialized by PowerSampleLogic.ps1.
# Do NOT reinitialize $script: variables here.

# ---------------------------
#region UI configuration
# ---------------------------

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

$script:CurrentView = "list"
$script:CurrentViewParam = 20
$script:CurrentTopList = @()
$script:MonitoringStartTime = Get-Date
$script:MeasurementInterval = 2

$script:GpuName = $null
$script:GpuIdlePowerMin = $null
$script:GpuIdlePowerMax = $null
$script:GpuIdleFanPercent = $null
$script:GpuIdleTemperatureC = $null

#endregion

# ---------------------------
#region UI helpers
# ---------------------------

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

# ---------------------------
#region Read helpers for CSV
# ---------------------------

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

# --------------------------
# GPU CSV readers (real-time)
# --------------------------

function Read-LatestGpuProcessAggregates {
    param(
        [string]$GpuProcessesCsvPath = $script:ProcessesCsvPath,
        [int]$TailLines = 4096
    )

    $result = @{}

    if (-not $GpuProcessesCsvPath) { return $result }
    if (-not (Test-Path $GpuProcessesCsvPath)) { return $result }

    # Read last N lines efficiently
    $lines = Get-Content -Path $GpuProcessesCsvPath -Tail $TailLines -ErrorAction SilentlyContinue
    if (-not $lines) { return $result }

    # Pattern for: Timestamp,PID,ProcessName,SMUtil,MemUtil,EncUtil,DecUtil,PowerW,EnergyJ,AccumulatedProcessEnergyJ,WeightedUtil,IsIdle
    $pattern = '^\s*(?:"(?<ts>[^"]*)"|(?<ts>[^,]*))\s*,\s*(?<pid>[^,]+)\s*,\s*(?:"(?<name>(?:[^"]|"")*)"|(?<name>[^,]*))\s*,\s*(?<sm>[^,]+)\s*,\s*(?<mem>[^,]+)\s*,\s*(?<enc>[^,]+)\s*,\s*(?<dec>[^,]+)\s*,\s*(?<pw>[^,]+)\s*,\s*(?<ej>[^,]+)\s*,\s*(?<acc>[^,]+)\s*,\s*(?<w>[^,]+)\s*,\s*(?<idle>[^,]+)\s*$'
    $regex = [System.Text.RegularExpressions.Regex]::new($pattern, [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ci = [System.Globalization.CultureInfo]::InvariantCulture

    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line.TrimStart().StartsWith('Timestamp', [System.StringComparison]::OrdinalIgnoreCase)) { continue }

        $m = $regex.Match($line)
        if (-not $m.Success) { continue }

        $pidRaw = $m.Groups['pid'].Value.Trim()
        # parse pid as int where possible
        $pidInt = 0
        try { [int]::TryParse($pidRaw, [ref]$pidInt) | Out-Null } catch { $pidInt = 0 }

        $procNameRaw = $m.Groups['name'].Value
        $procName = $procNameRaw -replace '""', '"'   # unescape CSV double-quotes

        # parse numeric fields defensively
        $sm = 0.0; $mem = 0.0; $enc = 0.0; $dec = 0.0
        $powerW = 0.0; $energyJ = 0.0; $accumJ = 0.0; $weighted = 0.0
        $isIdle = $false
        try { $sm = [double]::Parse($m.Groups['sm'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $mem = [double]::Parse($m.Groups['mem'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $enc = [double]::Parse($m.Groups['enc'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $dec = [double]::Parse($m.Groups['dec'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $powerW = [double]::Parse($m.Groups['pw'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $energyJ = [double]::Parse($m.Groups['ej'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $accumJ = [double]::Parse($m.Groups['acc'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $weighted = [double]::Parse($m.Groups['w'].Value, [System.Globalization.NumberStyles]::Float, $ci) } catch {}
        try { $isIdle = [bool]::Parse($m.Groups['idle'].Value) } catch {}

        # last occurrence wins because we're reading tail-to-head chronological order
        $key = if ($pidInt -ne 0) { $pidInt } else { $pidRaw }

        $result[$key] = @{
            PID = $key
            ProcessName = $procName
            SMUtil = $sm
            MemUtil = $mem
            EncUtil = $enc
            DecUtil = $dec
            PowerW = $powerW
            EnergyJ = $energyJ
            AccumulatedProcessEnergyJ = $accumJ
            WeightedUtil = $weighted
            IsIdle = $isIdle
        }
    }

    return $result
}

function Read-LatestGpuSample {
    param(
        [string]$GpuSamplesCsvPath = $script:SamplesCsvPath,
        [int]$TailLines = 20
    )

    if (-not $GpuSamplesCsvPath) { return $null }
    if (-not (Test-Path $GpuSamplesCsvPath)) { return $null }

    $lines = Get-Content -Path $GpuSamplesCsvPath -Tail $TailLines -ErrorAction SilentlyContinue
    if (-not $lines) { return $null }

    # Pattern:
    # Timestamp,PowerW,ActivePowerW,ExcessPowerW,GpuSMUtil,GpuMemUtil,GpuEncUtil,GpuDecUtil,GpuWeightTotal,ProcessWeightTotal,ProcessCount,AttributedPowerW,ResidualPowerW,AccumulatedEnergyJ,TemperatureC,FanPercent
    $pattern = '^\s*(?:"(?<ts>[^"]*)"|(?<ts>[^,]*))\s*,\s*(?<pw>[^,]+)\s*,\s*(?<active>[^,]+)\s*,\s*(?<excess>[^,]+)\s*,\s*(?<sm>[^,]+)\s*,\s*(?<mem>[^,]+)\s*,\s*(?<enc>[^,]+)\s*,\s*(?<dec>[^,]+)\s*,\s*(?<wtotal>[^,]+)\s*,\s*(?<pwtotal>[^,]+)\s*,\s*(?<pcnt>[^,]+)\s*,\s*(?<attr>[^,]+)\s*,\s*(?<res>[^,]+)\s*,\s*(?<acc>[^,]+)\s*,\s*(?<temp>[^,]+)\s*,\s*(?<fan>[^,]+)\s*$'
    $regex = [System.Text.RegularExpressions.Regex]::new($pattern, [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ci = [System.Globalization.CultureInfo]::InvariantCulture

    # pick last non-header matching line
    for ($i = $lines.Length - 1; $i -ge 0; $i--) {
        $line = $lines[$i]
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line.TrimStart().StartsWith('Timestamp', [System.StringComparison]::OrdinalIgnoreCase)) { continue }

        $m = $regex.Match($line)
        if (-not $m.Success) { continue }

        $ts = $m.Groups['ts'].Value
        $pw = 0.0; $active = 0.0; $excess = 0.0; $sm=0.0; $mem=0.0; $enc=0.0; $dec=0.0
        $wtotal = 0.0; $pwtotal = 0.0; $pcnt = 0; $attr = 0.0; $res = 0.0; $acc = 0.0; $temp=0.0; $fan=0.0
        try { $pw = [double]::Parse($m.Groups['pw'].Value, $ci) } catch {}
        try { $active = [double]::Parse($m.Groups['active'].Value, $ci) } catch {}
        try { $excess = [double]::Parse($m.Groups['excess'].Value, $ci) } catch {}
        try { $sm = [double]::Parse($m.Groups['sm'].Value, $ci) } catch {}
        try { $mem = [double]::Parse($m.Groups['mem'].Value, $ci) } catch {}
        try { $enc = [double]::Parse($m.Groups['enc'].Value, $ci) } catch {}
        try { $dec = [double]::Parse($m.Groups['dec'].Value, $ci) } catch {}
        try { $wtotal = [double]::Parse($m.Groups['wtotal'].Value, $ci) } catch {}
        try { $pwtotal = [double]::Parse($m.Groups['pwtotal'].Value, $ci) } catch {}
        try { $pcnt = [int]::Parse($m.Groups['pcnt'].Value) } catch {}
        try { $attr = [double]::Parse($m.Groups['attr'].Value, $ci) } catch {}
        try { $res = [double]::Parse($m.Groups['res'].Value, $ci) } catch {}
        try { $acc = [double]::Parse($m.Groups['acc'].Value, $ci) } catch {}
        try { $temp = [double]::Parse($m.Groups['temp'].Value, $ci) } catch {}
        try { $fan = [double]::Parse($m.Groups['fan'].Value, $ci) } catch {}

        return @{
            Timestamp = $ts
            PowerW = $pw
            ActivePowerW = $active
            ExcessPowerW = $excess
            GpuSmUtil = $sm
            GpuMemUtil = $mem
            GpuEncUtil = $enc
            GpuDecUtil = $dec
            GpuWeightTotal = $wtotal
            ProcessWeightTotal = $pwtotal
            ProcessCount = $pcnt
            AttributedPowerW = $attr
            ResidualPowerW = $res
            AccumulatedEnergyJ = $acc
            TemperatureC = $temp
            FanPercent = $fan
        }
    }

    return $null
}

#endregion

# ---------------------------
#region Main interactive loop
# ---------------------------

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

    # ---------------- GPU REAL-TIME SUMMARY (Report()-equivalent) ----------------
    # Determine GPU CSV paths (fallback to script globals)
    $gpuSamplesPath = if ($script:SamplesCsvPath) { $script:SamplesCsvPath } else { $null }
    $gpuProcsPath   = if ($script:ProcessesCsvPath) { $script:ProcessesCsvPath } else { $null }

    # Try to read GPU sample & per-process aggregates from CSV; fall back to in-memory if missing
    $gpuEntries = @{}   # pid -> aggregated info
    $latestGpuSample = $null

    if ($gpuProcsPath -and (Test-Path $gpuProcsPath)) {
        try {
            $gpuEntries = Read-LatestGpuProcessAggregates -GpuProcessesCsvPath $gpuProcsPath -TailLines 4096
        } catch { $gpuEntries = @{} }
    }

    if ($gpuSamplesPath -and (Test-Path $gpuSamplesPath)) {
        try {
            $latestGpuSample = Read-LatestGpuSample -GpuSamplesCsvPath $gpuSamplesPath -TailLines 20
        } catch { $latestGpuSample = $null }
    }

    # Build entriesList for display (prefer CSV aggregates; otherwise fallback to in-memory ProcessEnergyJoules)
    $entriesList = New-Object 'System.Collections.Generic.List[object]'

    if ($gpuEntries.Count -gt 0) {
        foreach ($kv in $gpuEntries.GetEnumerator()) {
            $pidKey = $kv.Key
            $acc = if ($kv.Value.AccumulatedProcessEnergyJ) { [double]$kv.Value.AccumulatedProcessEnergyJ } else { 0.0 }
            $entriesList.Add(@{ Key = $pidKey; Value = $acc; Name = $kv.Value.ProcessName })
        }
    } else {
        # fallback to in-memory map
        $iter = $script:ProcessEnergyJoules.GetEnumerator()
        while ($iter.MoveNext()) {
            $entriesList.Add(@{ Key = $iter.Current.Key; Value = [double]$iter.Current.Value })
        }
    }

    # Choose which total GPU energy to display: CSV sample preferred, then script value
    $gpuTotalEnergyToDisplay = if ($latestGpuSample -and ($null -ne $latestGpuSample.AccumulatedEnergyJ)) { $latestGpuSample.AccumulatedEnergyJ } else { $script:GpuEnergyJoules }

    # Map PIDs -> Names
    $allPids = $entriesList | ForEach-Object { [int]$_.Key } | Sort-Object -Unique
    $pidToName = @{}
    # First, use CSV-provided names (if present)
    if ($gpuEntries.Count -gt 0) {
        foreach ($kv in $gpuEntries.GetEnumerator()) {
            $procId = $kv.Key
            $pidToName[[int]$procId] = $kv.Value.ProcessName
        }
    }
    # Then try Get-Process for any missing names
    $missingPids = @()
    foreach ($procId in $allPids) { if (-not $pidToName.ContainsKey($procId)) { $missingPids += $procId } }
    if ($missingPids.Count -gt 0) {
        try {
            $procs = Get-Process -Id $missingPids -ErrorAction SilentlyContinue
            foreach ($pr in $procs) { $pidToName[[int]$pr.Id] = $pr.ProcessName }
        } catch {}
    }
    # Final fallback to samples array for names (already done in old logic)
    foreach ($procId in $allPids) {
        if (-not $pidToName.ContainsKey($procId)) {
            foreach ($s in $script:Samples) {
                if ($s.PerProcess -and $s.PerProcess.ContainsKey([string]$procId)) {
                    $pp = $s.PerProcess[[string]$procId]
                    if ($pp -and $pp.ProcessName) {
                        $pidToName[[int]$procId] = [string]$pp.ProcessName
                        break
                    }
                }
            }
        }
    }

    # Sum per-process energies
    $sumAttributedEnergy = 0.0
    foreach ($e in $entriesList) { $sumAttributedEnergy += [double]$e.Value }

    # Display GPU summary (uses gpuTotalEnergyToDisplay and entriesList)
    Write-Host "`n========================================================" -ForegroundColor $script:Colors.Header
    Write-Host "   GPU Process Power Monitor (Realtime)                  " -ForegroundColor $script:Colors.Header
    Write-Host "========================================================" -ForegroundColor $script:Colors.Header
    Write-Host ""
    Write-Host ("  Total GPU Energy (measured):   {0:F2} J" -f $gpuTotalEnergyToDisplay) -ForegroundColor Yellow
    Write-Host ("  Sum Attributed GPU Energy:     {0:F2} J" -f $sumAttributedEnergy) -ForegroundColor Yellow

    if ($sumAttributedEnergy -gt 0) {
        $diff = [math]::Abs($gpuTotalEnergyToDisplay - $sumAttributedEnergy)
        $pct  = if ($gpuTotalEnergyToDisplay -gt 0) { 100 * $diff / $gpuTotalEnergyToDisplay } else { 0 }
        Write-Host ("  Attribution Diff:              {0:F2} J ({1:F1}%)" -f $diff, $pct)
    } else {
        Write-Host ""
    }

    Write-Host ""
    Write-Host ("{0,-10} {1,-30} {2,12} {3,8}" -f "PID","Process","Energy (J)","GPU %") -ForegroundColor $script:Colors.Header
    Write-Host ("-" * 70)

    $entriesList |
        Sort-Object { [double]$_.'Value' } -Descending |
        ForEach-Object {
            $procId = [int]$_.Key
            $energy = [double]$_.Value
            $name = if ($pidToName.ContainsKey($procId)) { $pidToName[$procId] } else { (if ($_.Name) { $_.Name } else { "[exited]" }) }
            $pct = if ($sumAttributedEnergy -gt 0) { 100 * $energy / $sumAttributedEnergy } else { 0 }
            Write-Host ("{0,-10} {1,-30} {2,12:F2} {3,7:F1}%" -f $procId, $name, $energy, $pct)
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

function Start-CommandLineMode {
    param([int]$IntervalSeconds)

    # Get GPU name
    $gpuNameOutput = nvidia-smi --query-gpu=name --format=csv,noheader 2>$null
    $script:GpuName = if ($gpuNameOutput) { $gpuNameOutput.Trim() } else { "unknown" }

    Clear-Host
    Write-Host "`n========================================================" -ForegroundColor $script:Colors.Header
    Write-Host "   Interactive Process Power Monitor                    " -ForegroundColor $script:Colors.Header
    Write-Host "========================================================" -ForegroundColor $script:Colors.Header
    Write-Host ""
    Write-Host "CPU: $($script:CpuHardware.Name)" -ForegroundColor $script:Colors.Value
    Write-Host "GPU: $($script:GpuName)" -ForegroundColor $script:Colors.Value
    Write-Host "========================================================" -ForegroundColor $script:Colors.Header

    Write-Host ""
    Write-Host "Collecting initial CPU data..." -ForegroundColor $script:Colors.Info

    # Stopwatch
    if (-not $script:Stopwatch) {
        $script:Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    } else {
        $script:Stopwatch.Restart()
    }

    if (-not $script:StopwatchFrequency) {
        $script:StopwatchFrequency = [int64][System.Diagnostics.Stopwatch]::Frequency
    }

    # CPU warm-up
    $script:LastCpuSampleTicks = $script:Stopwatch.ElapsedTicks
    for ($i = 0; $i -lt 3; $i++) {
        Update-CPUEnergyData
        Start-Sleep -Seconds $IntervalSeconds
        Write-Host "." -NoNewline
    }

    # Measure idle GPU
    Write-Host "`nMeasuring idle GPU power..."
    Measure-IdleMetrics

    Write-Host "`nStarting interactive view...`n" -ForegroundColor Green
    Start-Sleep -Seconds 1

    $script:CurrentView = "list"
    $script:CurrentViewParam = 20
    [Console]::CursorVisible = $false

    # === GPU SAMPLING ===

    $intervalTicksGPU = [int64]([double]$script:SampleIntervalMs * $script:StopwatchFrequency / 1000.0)
    if ($intervalTicksGPU -le 0) { $intervalTicksGPU = 1 }
    $script:LastSampleTime = $script:Stopwatch.ElapsedTicks
    $scheduledTickGPU = $script:LastSampleTime + $intervalTicksGPU

    # CPU sampling
    $script:MeasurementInterval = $IntervalSeconds
    $intervalTicksCPU = [int64]([double]$script:MeasurementInterval * $script:StopwatchFrequency)
    if ($intervalTicksCPU -le 0) { $intervalTicksCPU = 1 }
    $scheduledTickCPU = $script:Stopwatch.ElapsedTicks + $intervalTicksCPU

    # Display refresh (2s)
    $displayIntervalTicks = [int64]([double]2000 * $script:StopwatchFrequency / 1000.0)
    $scheduledTickDisplay = $script:Stopwatch.ElapsedTicks + $displayIntervalTicks

    $sampleCount = 0

    try {
        Show-TopList -Count $script:CurrentViewParam -ClearScreen $true
        [Console]::CursorVisible = $true

        while ($true) {
            $now = $script:Stopwatch.ElapsedTicks

            # === GPU sampling (bit-for-bit old behavior) ===
            $remainingGPU = $scheduledTickGPU - $now
            if ($remainingGPU -le 0) {
                Update-GPUEnergyData
                $sampleCount++

                do {
                    $scheduledTickGPU += $intervalTicksGPU
                    $now = $script:Stopwatch.ElapsedTicks
                } while ($scheduledTickGPU -le $now)
            }

            # CPU sampling
            if (($scheduledTickCPU - $now) -le 0) {
                Update-CPUEnergyData
                do {
                    $scheduledTickCPU += $intervalTicksCPU
                    $now = $script:Stopwatch.ElapsedTicks
                } while ($scheduledTickCPU -le $now)
            }

            # Display refresh
            if (($scheduledTickDisplay - $now) -le 0) {
                [Console]::CursorVisible = $false
                switch ($script:CurrentView) {
                    "list"  { Show-TopList -Count $script:CurrentViewParam -ClearScreen $false -SampleCount $sampleCount }
                    "focus" { Show-FocusedView -ProcessName $script:CurrentViewParam }
                }
                [Console]::CursorVisible = $true

                do {
                    $scheduledTickDisplay += $displayIntervalTicks
                    $now = $script:Stopwatch.ElapsedTicks
                } while ($scheduledTickDisplay -le $now)
            }

            # Sleep logic
            $nextTick = @($scheduledTickGPU, $scheduledTickCPU, $scheduledTickDisplay | Where-Object { $_ -gt $now } | Measure-Object -Minimum).Minimum
            if ($nextTick) {
                $remainingMs = [int]([double]($nextTick - $now) * 1000.0 / $script:StopwatchFrequency)
                if ($remainingMs -gt 5) {
                    Start-Sleep -Milliseconds ($remainingMs - 3)
                } else {
                    [System.Threading.Thread]::Sleep(0)
                }
            }

            # Check for user input (non-blocking)
            if ([Console]::KeyAvailable) {
                $inputLine = [Console]::ReadLine()
                if ($inputLine -match '^(quit|exit)$') { throw "EXIT" }


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
        try { Close-Writers } catch {}
        [Console]::CursorVisible = $true
        if ($_.Exception.Message -eq "EXIT") {
            Write-Host "`nExiting..." -ForegroundColor $script:Colors.Info
        }
    }
}

#endregion