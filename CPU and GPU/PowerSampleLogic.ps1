# ===========================
# PowerSampleLogic.ps1
# Core sampling + attribution logic (PRODUCER)
# ===========================

# ---------------------------
#region Script-scoped state
# ---------------------------

$script:Samples = New-Object System.Collections.Generic.List[object]
$script:ProcessEnergyJoules = @{}
$script:GpuIdleProcesses = @{}
$script:ProcessEnergyData = @{}

$script:GpuEnergyJoules = 0.0
$script:GpuIdlePower = 0.0
$script:IdleProcessCount = 0

$script:SamplesCsvWriter = $null
$script:ProcessesCsvWriter = $null
$script:CpuSamplesCsvWriter = $null
$script:CpuProcessesCsvWriter = $null

$script:SamplesCsvPath = $null
$script:ProcessesCsvPath = $null
$script:CpuSamplesCsvPath = $null
$script:CpuProcessesCsvPath = $null

$script:SamplesSinceLastFlush = 0
$script:SamplesFlushInterval = 10
$script:MaxSamplesInMemory = 2000
$script:MaxHistorySize = 120

$script:WeightSM = 1.0
$script:WeightMemory = 0.5
$script:WeightEncoder = 0.25
$script:WeightDecoder = 0.15

$script:SampleIntervalMs = 100
$script:DiagnosticsOutputPath = $null
$script:TimeStampLogging = Get-Date

$script:Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$script:StopwatchFrequency = [double][System.Diagnostics.Stopwatch]::Frequency
$script:LastSampleTime = 0
$script:LastCpuSampleTicks = 0

$script:CurrentCpuPower = 0.0
$script:CurrentSystemPower = 0.0
$script:CurrentTotalCpuPercent = 0.0
$script:MeasurementCount = 0

$script:ProcessNameCache = $null
$script:ProcessNameCacheOrder = $null
$script:ProcessNameCacheCapacity = $1024

#endregion

# ---------------------------
#region Initialization
# ---------------------------

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

function Initialize-GPUProcessMonitor {
    param(
        [int]$SampleIntervalMs,
        [double]$wSM,
        [double]$wMem,
        [double]$wEnc,
        [double]$wDec,
        [string]$diagPath
    )

    $script:SampleIntervalMs = $SampleIntervalMs
    $script:WeightSM = $wSM
    $script:WeightMemory = $wMem
    $script:WeightEncoder = $wEnc
    $script:WeightDecoder = $wDec
    $script:DiagnosticsOutputPath = $diagPath
    $script:TimeStampLogging = Get-Date

    Open-Writers
}

#endregion

# ---------------------------
#region Writers
# ---------------------------

 function Open-Writers {
    # Setup diagnostics paths and open writers immediately (header creation happens here)
    try {
        $timestampForFile = $script:TimeStampLogging.ToString('yyyyMMdd_HHmmss')
        $outSpec = $script:DiagnosticsOutputPath
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

        # Persist paths for consumer UI:
        $script:SamplesCsvPath = $csvPath
        $script:ProcessesCsvPath = $csvPathProcesses

        # --- GPU sample writer ---
        $needHeaderSamples = -not (Test-Path $csvPath)
        $fsSamples = [System.IO.FileStream]::new($csvPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
        $fsSamples.Seek(0, [System.IO.SeekOrigin]::End) | Out-Null
        $script:SamplesCsvWriter = [System.IO.StreamWriter]::new($fsSamples, [System.Text.Encoding]::UTF8)
        $script:SamplesCsvWriter.AutoFlush = $false
        if ($needHeaderSamples -and ($fsSamples.Length -eq 0)) {
            $script:SamplesCsvWriter.WriteLine('Timestamp,PowerW,ActivePowerW,ExcessPowerW,GpuSMUtil,GpuMemUtil,GpuEncUtil,GpuDecUtil,GpuWeightTotal,ProcessWeightTotal,ProcessCount,AttributedPowerW,ResidualPowerW,AccumulatedEnergyJ,TemperatureC,FanPercent')
            $script:SamplesCsvWriter.Flush()
        }

        # --- GPU per-process writer ---
        $needHeaderProcs = -not (Test-Path $csvPathProcesses)
        $fsProcs = [System.IO.FileStream]::new($csvPathProcesses, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
        $fsProcs.Seek(0, [System.IO.SeekOrigin]::End) | Out-Null
        $script:ProcessesCsvWriter = [System.IO.StreamWriter]::new($fsProcs, [System.Text.Encoding]::UTF8)
        $script:ProcessesCsvWriter.AutoFlush = $false
        if ($needHeaderProcs -and ($fsProcs.Length -eq 0)) {
            $script:ProcessesCsvWriter.WriteLine('Timestamp,PID,ProcessName,SMUtil,MemUtil,EncUtil,DecUtil,PowerW,EnergyJ,AccumulatedEnergyJ,WeightedUtil,IsIdle')
            $script:ProcessesCsvWriter.Flush()
        }

        # --- CPU writers ---
        $cpuSamplesPath = if ($outSpec -match '\.csv$') { $outSpec -replace '\.csv$','_cpu_samples.csv' } else { Join-Path $diagDir ("cpu_samples_$timestampForFile.csv") }
        $cpuProcsPath   = if ($outSpec -match '\.csv$') { $outSpec -replace '\.csv$','_cpu_processes.csv' } else { Join-Path $diagDir ("cpu_processes_$timestampForFile.csv") }

        # Persist CPU paths for consumer UI:
        $script:CpuSamplesCsvPath = $cpuSamplesPath
        $script:CpuProcessesCsvPath = $cpuProcsPath

        # CPU samples writer
        $needHeaderCpuSamples = -not (Test-Path $cpuSamplesPath)
        $fsCpuSamples = [System.IO.FileStream]::new($cpuSamplesPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
        $fsCpuSamples.Seek(0, [System.IO.SeekOrigin]::End) | Out-Null
        $script:CpuSamplesCsvWriter = [System.IO.StreamWriter]::new($fsCpuSamples, [System.Text.Encoding]::UTF8)
        $script:CpuSamplesCsvWriter.AutoFlush = $false
        if ($needHeaderCpuSamples -and ($fsCpuSamples.Length -eq 0)) {
            $script:CpuSamplesCsvWriter.WriteLine('Timestamp,CpuPowerMw,SystemPowerMw,TotalCpuPercent,MeasurementIntervalSeconds,AccumulatedCpuEnergymJ')
            $script:CpuSamplesCsvWriter.Flush()
        }

        # CPU per-process writer
        $needHeaderCpuProcs = -not (Test-Path $cpuProcsPath)
        $fsCpuProcs = [System.IO.FileStream]::new($cpuProcsPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
        $fsCpuProcs.Seek(0, [System.IO.SeekOrigin]::End) | Out-Null
        $script:CpuProcessesCsvWriter = [System.IO.StreamWriter]::new($fsCpuProcs, [System.Text.Encoding]::UTF8)
        $script:CpuProcessesCsvWriter.AutoFlush = $false
        if ($needHeaderCpuProcs -and ($fsCpuProcs.Length -eq 0)) {
            $script:CpuProcessesCsvWriter.WriteLine('Timestamp,ProcessName,CpuPercent,ProcessPowerMw,EnergymJ,AccumulatedEnergymJ')
            $script:CpuProcessesCsvWriter.Flush()
        }
    } catch {
        Write-Host "Warning: Could not create diagnostics files: $($_.Exception.Message)" -ForegroundColor Yellow

        # Clean up any writers that were partially created to avoid leaked handles
        try { if ($null -ne $script:SamplesCsvWriter) { try { $script:SamplesCsvWriter.Flush() } catch {}; try { $script:SamplesCsvWriter.Close() } catch {}; $script:SamplesCsvWriter = $null } } catch {}
        try { if ($null -ne $script:ProcessesCsvWriter) { try { $script:ProcessesCsvWriter.Flush() } catch {}; try { $script:ProcessesCsvWriter.Close() } catch {}; $script:ProcessesCsvWriter = $null } } catch {}
        try { if ($null -ne $script:CpuSamplesCsvWriter) { try { $script:CpuSamplesCsvWriter.Flush() } catch {}; try { $script:CpuSamplesCsvWriter.Close() } catch {}; $script:CpuSamplesCsvWriter = $null } } catch {}
        try { if ($null -ne $script:CpuProcessesCsvWriter) { try { $script:CpuProcessesCsvWriter.Flush() } catch {}; try { $script:CpuProcessesCsvWriter.Close() } catch {}; $script:CpuProcessesCsvWriter = $null } } catch {}
    }
}

# helper to safely flush+close writers (called on exit)
function Close-Writers {
    try {
        if ($null -ne $script:SamplesCsvWriter) {
            try { $script:SamplesCsvWriter.Flush() } catch {}
            try { $script:SamplesCsvWriter.Close() } catch {}
            $script:SamplesCsvWriter = $null
        }
    } catch {}

    try {
        if ($null -ne $script:ProcessesCsvWriter) {
            try { $script:ProcessesCsvWriter.Flush() } catch {}
            try { $script:ProcessesCsvWriter.Close() } catch {}
            $script:ProcessesCsvWriter = $null
        }
    } catch {}

    try {
        if ($null -ne $script:CpuSamplesCsvWriter) {
            try { $script:CpuSamplesCsvWriter.Flush() } catch {}
            try { $script:CpuSamplesCsvWriter.Close() } catch {}
            $script:CpuSamplesCsvWriter = $null
        }
    } catch {}

    try {
        if ($null -ne $script:CpuProcessesCsvWriter) {
            try { $script:CpuProcessesCsvWriter.Flush() } catch {}
            try { $script:CpuProcessesCsvWriter.Close() } catch {}
            $script:CpuProcessesCsvWriter = $null
        }
    } catch {}
}

#endregion

# ---------------------------
#region Utility
# ---------------------------

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

function Parse-Util {
    param([string]$value)

    if ([string]::IsNullOrWhiteSpace($value) -or $value -eq '-') { return 0.0 }

    $s = $value.Trim()
    $s = $s -replace '[^\d\.\-+]',''

    $d = 0.0
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
        return $d
    }

    return 0.0
}

#endregion

# ---------------------------
#region CPU sampling 
# ---------------------------

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

function Update-CPUEnergyData {
    # --- timing (same model as GPU) ---
    $currentTicks = $script:Stopwatch.ElapsedTicks

    if ($script:LastCpuSampleTicks -eq 0) {
        # first call: initialize timing baseline only (no energy can be computed yet)
        $script:LastCpuSampleTicks = $currentTicks
        return $false
    }

    $deltaTicks = $currentTicks - $script:LastCpuSampleTicks
    if ($deltaTicks -le 0) { return $false }

    $dt = $deltaTicks / $script:StopwatchFrequency    # seconds (double)
    $script:LastCpuSampleTicks = $currentTicks

    # --- acquire measurements ---
    $cpuData = Get-ProcessCpuUtilization
    $cpuPowerMilliwatts = Get-CpuPowerConsumption
    $systemPowerMilliwatts = Get-SystemPowerConsumption

    if ($null -eq $cpuData -or $null -eq $cpuPowerMilliwatts) {
        return $false
    }

    # --- update globals ---
    $script:CurrentCpuPower = $cpuPowerMilliwatts
    $script:CurrentSystemPower = if ($null -ne $systemPowerMilliwatts) { $systemPowerMilliwatts } else { 0 }
    $script:CurrentTotalCpuPercent = $cpuData.TotalCpu
    $script:MeasurementCount++

    $totalCpu = $cpuData.TotalCpu
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $timestampIso = (Get-Date).ToString("o")
    $dtSeconds = [double]$dt

    # If there are process contributions, update them; otherwise continue so we still log the sample
    if ($totalCpu -gt 0) {
        foreach ($processEntry in $cpuData.ProcessData.GetEnumerator()) {

            $processName = $processEntry.Key
            $processCpu  = [double]$processEntry.Value
            $cpuRatio    = $processCpu / $totalCpu

            $processPowerCpu = $cpuPowerMilliwatts * $cpuRatio
            $energyCpu = $processPowerCpu * $dtSeconds   # ← stopwatch-driven (mW * s = mJ)

            $energySystem = 0.0
            if ($null -ne $systemPowerMilliwatts) {
                $energySystem = ($systemPowerMilliwatts * $cpuRatio) * $dtSeconds
            }

            if (-not $script:ProcessEnergyData.ContainsKey($processName)) {
                $script:ProcessEnergyData[$processName] = @{
                    CpuEnergy        = 0.0
                    SystemEnergy     = 0.0
                    LastSeenCpu      = 0.0
                    LastSeenPowerMw  = 0.0
                    PowerHistory     = New-Object System.Collections.ArrayList
                }
            }

            $entry = $script:ProcessEnergyData[$processName]
            $entry.CpuEnergy       += $energyCpu
            $entry.SystemEnergy    += $energySystem
            $entry.LastSeenCpu      = $processCpu
            $entry.LastSeenPowerMw  = $processPowerCpu

            if ($entry.PowerHistory.Count -ge $script:MaxHistorySize) {
                $entry.PowerHistory.RemoveAt(0)
            }
            [void]$entry.PowerHistory.Add($processPowerCpu)
        }
    }

    # ---------------- CSV logging ----------------
    try {
        # CPU samples (overall)
        if ($null -ne $script:CpuSamplesCsvWriter) {
            $accCpuEnergy = 0.0
            foreach ($value in $script:ProcessEnergyData.Values) {
                $accCpuEnergy += [double]$value.CpuEnergy
            }

            $line = '"' + $timestampIso + '",' +
                    ([double]$cpuPowerMilliwatts).ToString('F4', $inv) + ',' +
                    ([double]$script:CurrentSystemPower).ToString('F4', $inv) + ',' +
                    ([double]$script:CurrentTotalCpuPercent).ToString('F2', $inv) + ',' +
                    $dtSeconds.ToString('F6', $inv) + ',' +
                    $accCpuEnergy.ToString('F4', $inv)

            $script:CpuSamplesCsvWriter.WriteLine($line)
        }

        # CPU per-process
        if ($null -ne $script:CpuProcessesCsvWriter) {
            foreach ($procEntry in $script:ProcessEnergyData.GetEnumerator()) {
                $processName = $procEntry.Key
                $data = $procEntry.Value

                # use measured dt (stopwatch) for per-process sample energy
                $energyThisSample = [double]$data.LastSeenPowerMw * $dtSeconds

                $escapedName = ($processName -replace '"','""') -replace '\r|\n',' '

                $line2 = '"' + $timestampIso + '","' + $escapedName + '",' +
                        ([double]$data.LastSeenCpu).ToString('F4', $inv) + ',' +
                        ([double]$data.LastSeenPowerMw).ToString('F4', $inv) + ',' +
                        $energyThisSample.ToString('F4', $inv) + ',' +
                        ([double]$data.CpuEnergy).ToString('F4', $inv)

                $script:CpuProcessesCsvWriter.WriteLine($line2)
            }
        }

        # batched flush (shared counter with GPU path)
        $script:SamplesSinceLastFlush += 1
        if ($script:SamplesSinceLastFlush -ge $script:SamplesFlushInterval) {
            try {
                if ($null -ne $script:CpuSamplesCsvWriter)   { $script:CpuSamplesCsvWriter.Flush() }
                if ($null -ne $script:CpuProcessesCsvWriter) { $script:CpuProcessesCsvWriter.Flush() }
            } catch {}
            $script:SamplesSinceLastFlush = 0
        }
    }
    catch {
        Write-Host "Realtime CPU CSV logging failed: $($_.Exception.Message)"
    }

    return $true
}

#endregion

# ---------------------------
#region GPU sampling
# ---------------------------

function Measure-IdleMetrics {
    # local arrays for idle sampling
    $idleTemperatureSamples = @()
    $idleFanUtilSamples = @()
    $idlePowerSamples = @()
    $idlePowerMin = 0
    $idlePowerMax = 0
    for ($i = 0; $i -lt 50; $i++) {
        $out = nvidia-smi --query-gpu=power.draw.instant,temperature.gpu,fan.speed --format=csv,noheader,nounits 2>$null
        if ($out -and -not [string]::IsNullOrWhiteSpace($out) -and $out.Trim() -ne '[N/A]') {
            $parts = $out -split ',\s*'
            if ($parts.Count -ge 3) {
                try {
                    $idlePowerSamples += [double]$parts[0].Trim()
                    $idleTemperatureSamples += [double]$parts[1].Trim()
                    $idleFanUtilSamples += [double]$parts[2].Trim()
                }
                catch {}
            }
        }
        Start-Sleep -Milliseconds $script:SampleIntervalMs
    }

    if ($idleTemperatureSamples.Count -gt 0) { $script:GpuIdleTemperatureC = ($idleTemperatureSamples | Measure-Object -Average).Average }
    if ($idleFanUtilSamples.Count -gt 0) { $script:GpuIdleFanPercent = ($idleFanUtilSamples | Measure-Object -Average).Average }
    if ($idlePowerSamples.Count -gt 0) {
        $measuredIdlePower = ($idlePowerSamples | Measure-Object -Average).Average
        $idlePowerMin = $idlePowerSamples | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
        $idlePowerMax = $idlePowerSamples | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
        $script:GpuIdlePower = $measuredIdlePower
        $script:GpuIdlePowerMin = $idlePowerMin
        $script:GpuIdlePowerMax = $idlePowerMax
    }

    # Record which processes were running during idle measurement
    $idleProcesses = Get-GpuProcesses
    $script:IdleProcessCount = $idleProcesses.Count
    foreach ($idleProcessObj in $idleProcesses) {
        $script:GpuIdleProcesses[[string]$idleProcessObj.ProcessId] = $true
    }
}

function Get-GpuProcesses {
    # Initialize cache if needed
    if ($null -eq $script:ProcessNameCache) {
        $script:ProcessNameCache = [System.Collections.Generic.Dictionary[int,string]]::new()
        $script:ProcessNameCacheOrder = [System.Collections.Generic.LinkedList[int]]::new()
        if ($script:ProcessNameCacheCapacity -le 0) { $script:ProcessNameCacheCapacity = 1024 }
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

        $m = $pmonRegex.Match($line)
        if (-not $m.Success) { continue }

        $procIdInt = [int]$m.Groups[2].Value

        [void]$pidHash.Add($procIdInt)

        $rawEntries.Add([PSCustomObject]@{
            ProcessId = $procIdInt
            SmUtil    = (Parse-Util -value $m.Groups[4].Value)
            MemUtil   = (Parse-Util -value $m.Groups[5].Value)
            EncUtil   = (Parse-Util -value $m.Groups[6].Value)
            DecUtil   = (Parse-Util -value $m.Groups[7].Value)
            Command   = $m.Groups[10].Value.Trim()
        })
    }

    if ($rawEntries.Count -eq 0) { return $resultList.ToArray() }

    # local references for performance
    $processNameCacheLocal = $script:ProcessNameCache
    $processNameCacheOrderLocal = $script:ProcessNameCacheOrder
    $processNameCacheCapacityLocal = $script:ProcessNameCacheCapacity

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
            foreach ($wp in $winProcs) {
                $wpId = [int]$wp.Id
                $wpName = $wp.ProcessName
                if (-not $processNameCacheLocal.ContainsKey($wpId)) {
                    $processNameCacheLocal.Add($wpId, $wpName)
                    $processNameCacheOrderLocal.AddFirst($wpId)
                    while ($processNameCacheOrderLocal.Count -gt $processNameCacheCapacityLocal) {
                        $lastNode = $processNameCacheOrderLocal.Last
                        if ($lastNode) {
                            $oldPid = [int]$lastNode.Value
                            $processNameCacheOrderLocal.RemoveLast()
                            if ($processNameCacheLocal.ContainsKey($oldPid)) { $processNameCacheLocal.Remove($oldPid) }
                        } else { break }
                    }
                } else {
                    $existingNode = $processNameCacheOrderLocal.Find($wpId)
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

function Update-GPUEnergyData {
    # fast local caches of instance fields (avoid repeated $script: property lookup)
    $weightSmLocal      = $script:WeightSM
    $weightMemoryLocal  = $script:WeightMemory
    $weightEncoderLocal = $script:WeightEncoder
    $weightDecoderLocal = $script:WeightDecoder
    $gpuIdleProcessMap  = $script:GpuIdleProcesses
    $processEnergyMap   = $script:ProcessEnergyJoules

    # timing (hot path)
    $currentTicks = $script:Stopwatch.ElapsedTicks
    $deltaTicks   = $currentTicks - $script:LastSampleTime
    $script:LastSampleTime = $currentTicks
    if ($deltaTicks -le 0) { return }
    $dt = $deltaTicks / $script:StopwatchFrequency

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
    $processEntries = Get-GpuProcesses
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
    $gpuActivePower = if (($script:GpuIdlePower) -lt $gpuPower) { $gpuPower - ($script:GpuIdlePower) } else { 0 }
    $gpuTotalProcessUtil = if ($gpuWeightedTotal -gt 0) { $processWeightTotal / $gpuWeightedTotal } else { 0 }
    $gpuTotalActiveProcessPower = $gpuTotalProcessUtil * $gpuActivePower
    $gpuExcessPower = $gpuActivePower - $gpuTotalActiveProcessPower

    $activeProcessCount = if ($inactiveIdleProcessesCount -eq $currentProcessCount) { $currentProcessCount } else { $currentProcessCount - $inactiveIdleProcessesCount }
    $P_idle_pwr = if ($script:IdleProcessCount -gt 0) { $script:GpuIdlePower / $script:IdleProcessCount } else { 0 }
    $P_residual_per_proc = if ($activeProcessCount -gt 0) { $gpuExcessPower / $activeProcessCount } else { 0 }

    # per-process attribution objects
    $sampleAttributedPower = 0.0
    $samplePerProcess = @{ }

    # Precompute culture once
    $inv = [System.Globalization.CultureInfo]::InvariantCulture

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

        $energyJ = $powerValue * $dt

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
    $script:GpuEnergyJoules += ($gpuPower * $dt)
    $accumulatedEnergy = $script:GpuEnergyJoules
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

    [void]$script:Samples.Add($sampleObj)

    # enforce in-memory cap
    try {
        if ($script:Samples.Count -gt $script:MaxSamplesInMemory) {
            $removeCount = $script:Samples.Count - $script:MaxSamplesInMemory
            # RemoveRange exists on List<T>
            $script:Samples.RemoveRange(0, $removeCount)
        }
    } catch {}

    # ---------- realtime CSV logging (cached writers & batched flush) ----------
    try {
        # write sample line (single concatenated string) — round here only
        $line = '"' + $tsEsc_local + '",' +
                ([double]$gpuPower).ToString('F4', $inv) + ',' +
                ([double]$gpuActivePower).ToString('F4', $inv) + ',' +
                ([double]$gpuExcessPower).ToString('F4', $inv) + ',' +
                ([double]$gpuSmUtil).ToString('F2', $inv) + ',' +
                ([double]$gpuMemUtil).ToString('F2', $inv) + ',' +
                ([double]$gpuEncUtil).ToString('F2', $inv) + ',' +
                ([double]$gpuDecUtil).ToString('F2', $inv) + ',' +
                ([double]$gpuWeightedTotal).ToString('F2', $inv) + ',' +
                ([double]$processWeightTotal).ToString('F2', $inv) + ',' +
                [int]$currentProcessCount + ',' +
                ([double]$sampleAttributedPower).ToString('F6', $inv) + ',' +
                ([double]$sampleResidualPower).ToString('F6', $inv) + ',' +
                ([double]$accumulatedEnergy).ToString('F8', $inv) + ',' +
                ([double]$gpuTemperature).ToString('F1', $inv) + ',' +
                ([double]$gpuFanUtil).ToString('F1', $inv)

        $script:SamplesCsvWriter.WriteLine($line)

        # write per-process lines (round here)
        foreach ($perProcessEntry in $samplePerProcess.Values) {
            $pnameEsc = ($perProcessEntry.ProcessName -replace '"','""') -replace '\r|\n',' '
            $line2 = '"' + $tsEsc_local + '",' +
                    $perProcessEntry.PID + ',"' + $pnameEsc + '",' +
                    ([double]$perProcessEntry.SMUtil).ToString('F2', $inv) + ',' +
                    ([double]$perProcessEntry.MemUtil).ToString('F2', $inv) + ',' +
                    ([double]$perProcessEntry.EncUtil).ToString('F2', $inv) + ',' +
                    ([double]$perProcessEntry.DecUtil).ToString('F2', $inv) + ',' +
                    ([double]$perProcessEntry.PowerW).ToString('F6', $inv) + ',' +
                    ([double]$perProcessEntry.EnergyJ).ToString('F8', $inv) + ',' +
                    ([double]$perProcessEntry.AccumulatedProcessEnergyJ).ToString('F8', $inv) + ',' +
                    ([double]$perProcessEntry.WeightedUtil).ToString('F4', $inv) + ',' +
                    ([bool]$perProcessEntry.IsIdle).ToString()
            $script:ProcessesCsvWriter.WriteLine($line2)
        }

        # batched flush (reduce I/O churn)
        $script:SamplesSinceLastFlush += 1
        if ($script:SamplesSinceLastFlush -ge $script:SamplesFlushInterval) {
            try {
                $script:SamplesCsvWriter.Flush()
                $script:ProcessesCsvWriter.Flush()
            } catch {}
            $script:SamplesSinceLastFlush = 0
        }
    }
    catch {
        # do not break sampling
        Write-Host "Realtime CSV logging failed: $($_.Exception.Message)"
    }
}

#endregion

# regions are cool :), hope u can see them
# yeah I see them, I didn't even know it was a thing until now (〃￣︶￣)人(￣︶￣〃)