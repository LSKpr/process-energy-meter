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

<#
Flags of interest:
    -SampleInterval: how often to sample in ms (default 100)
    -WeightSM, WeightMem, WeightEnc, WeightDec: weights for the attribution formula (defaults: 1.0, 0.5, 0.25, 0.15)
    -DiagnosticsOutput: base path for writing diagnostics CSV files (default "power_diagnostics")
#>
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

$script:SampleIntervalMs = $null
$script:WeightSM = $null
$script:WeightMemory = $null
$script:WeightEncoder = $null
$script:WeightDecoder = $null
$script:Samples = $null
$script:Stopwatch = $null
$script:StopwatchFrequency = $null
$script:LastSampleTime = $null
$script:TimeStampLogging = $null
$script:ProcessEnergyJoules = $null
$script:GpuName = $null
$script:GpuEnergyJoules = $null
$script:GpuIdlePower = $null
$script:GpuIdlePowerMin = $null
$script:GpuIdlePowerMax = $null
$script:GpuIdleFanPercent = $null
$script:GpuIdleTemperatureC = $null
$script:GpuIdleProcesses = $null
$script:GpuCurrentPower = $null
$script:IdleProcessCount = $null
$script:DiagnosticsOutputPath = $null
$script:ProcessNameCache = $null
$script:ProcessNameCacheOrder = $null
$script:ProcessNameCacheCapacity = $null
$script:SamplesCsvWriter = $null
$script:ProcessesCsvWriter = $null
$script:SamplesSinceLastFlush = $null
$script:SamplesFlushInterval = $null
$script:MaxSamplesInMemory = $null

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
    $script:Samples = [System.Collections.Generic.List[object]]::new()
    $script:Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $script:StopwatchFrequency = [System.Diagnostics.Stopwatch]::Frequency
    $script:LastSampleTime = $script:Stopwatch.ElapsedTicks
    $script:TimeStampLogging = Get-Date
    $script:ProcessEnergyJoules = @{}
    $script:GpuEnergyJoules = 0.0
    $script:GpuIdleProcesses = @{}
    $script:DiagnosticsOutputPath = $diagPath
    $script:GpuCurrentPower = 0.0

    # initialize LRU cache defaults
    $script:ProcessNameCache = $null
    $script:ProcessNameCacheOrder = $null
    $script:ProcessNameCacheCapacity = 1024       # default capacity

    # writer flush tuning
    $script:SamplesSinceLastFlush = 0
    $script:SamplesFlushInterval = 10   # flush every 10 samples

    # long-run safety: cap in-memory samples
    $script:MaxSamplesInMemory = 10000  # default; adjust if you want to keep more samples in memory

    # Get GPU name
    $gpuNameOutput = nvidia-smi --query-gpu=name --format=csv,noheader 2>$null
    $script:GpuName = if ($gpuNameOutput) { $gpuNameOutput.Trim() } else { "unknown" }

    # Open Writers for diagnostics output (creates files and writes headers)
    Open-Writers
}

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
        Write-Host "." -NoNewline -ForegroundColor $script:Colors.Info
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

        # Use FileStream with FileShare.Read so other tools can read while we append (robust for long runs)
        $needHeaderSamples = -not (Test-Path $csvPath)
        $fsSamples = [System.IO.FileStream]::new($csvPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
        # Seek to end for append
        $fsSamples.Seek(0, [System.IO.SeekOrigin]::End) | Out-Null
        $script:SamplesCsvWriter = [System.IO.StreamWriter]::new($fsSamples, [System.Text.Encoding]::UTF8)
        $script:SamplesCsvWriter.AutoFlush = $false
        if ($needHeaderSamples -and ($fsSamples.Length -eq 0)) {
            $script:SamplesCsvWriter.WriteLine('Timestamp,PowerW,ActivePowerW,ExcessPowerW,GpuSMUtil,GpuMemUtil,GpuEncUtil,GpuDecUtil,GpuWeightTotal,ProcessWeightTotal,ProcessCount,AttributedPowerW,ResidualPowerW,AccumulatedEnergyJ,TemperatureC,FanPercent')
            $script:SamplesCsvWriter.Flush()
        }

        $needHeaderProcs = -not (Test-Path $csvPathProcesses)
        $fsProcs = [System.IO.FileStream]::new($csvPathProcesses, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
        $fsProcs.Seek(0, [System.IO.SeekOrigin]::End) | Out-Null
        $script:ProcessesCsvWriter = [System.IO.StreamWriter]::new($fsProcs, [System.Text.Encoding]::UTF8)
        $script:ProcessesCsvWriter.AutoFlush = $false
        if ($needHeaderProcs -and ($fsProcs.Length -eq 0)) {
            $script:ProcessesCsvWriter.WriteLine('Timestamp,PID,ProcessName,SMUtil,MemUtil,EncUtil,DecUtil,PowerW,EnergyJ,AccumulatedEnergyJ,WeightedUtil,IsIdle')
            $script:ProcessesCsvWriter.Flush()
        }

        # --- CPU CSV writers (add inside Open-Writers, alongside GPU writers setup) ---

        # choose base dir & timestamps consistent with existing logic
        $cpuSamplesPath = if ($outSpec -match '\.csv$') { $outSpec -replace '\.csv$','_cpu_samples.csv' } else { Join-Path $diagDir ("cpu_samples_$timestampForFile.csv") }
        $cpuProcsPath   = if ($outSpec -match '\.csv$') { $outSpec -replace '\.csv$','_cpu_processes.csv' } else { Join-Path $diagDir ("cpu_processes_$timestampForFile.csv") }

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
        Write-Host "Warning: Could not create diagnostics files in constructor: $($_.Exception.Message)"
        $script:SamplesCsvWriter = $null
        $script:ProcessesCsvWriter = $null
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
}


#endregion


#region Core Monitoring Logic

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

    try {
        # ensure writers exist (lazy open not shown here; rely on Open-Writers)
        if ($null -ne $script:CpuSamplesCsvWriter) {
            $ts = (Get-Date).ToString("o")
            $cpuPowerMw = [double]$cpuPowerMilliwatts
            $sysPowerMw = if ($script:CurrentSystemPower) { [double]$script:CurrentSystemPower } else { 0.0 }
            $totalCpuPct = [double]$script:CurrentTotalCpuPercent
            $intervalSec = [double]$IntervalSeconds
            # compute accumulated CPU energy (sum of all process CpuEnergy? you already store per-process cumulative values)
            # pick a representative accumulated CPU energy total (sum across ProcessEnergyData CpuEnergy)
            $accCpuEnergy = 0.0
            foreach ($v in $script:ProcessEnergyData.Values) { $accCpuEnergy += [double]$v.CpuEnergy }

            $line = "{0},{1:F4},{2:F4},{3:F2},{4:F3},{5:F4}" -f $ts, $cpuPowerMw, $sysPowerMw, $totalCpuPct, $intervalSec, $accCpuEnergy
            $script:CpuSamplesCsvWriter.WriteLine($line)
        }

        if ($null -ne $script:CpuProcessesCsvWriter) {
            $ts = (Get-Date).ToString("o")
            foreach ($procEntry in $script:ProcessEnergyData.GetEnumerator()) {
                $pname = $procEntry.Key
                $pdata = $procEntry.Value
                $procCpuPct = [double]$pdata.LastSeenCpu
                $procPowerMw = [double]$pdata.LastSeenPowerMw
                # energy this sample (we computed energyCpu above per-process as $energyCpu)
                # but we don't have a separate variable per loop here — compute sample energy:
                $energyThisSample = $procPowerMw * [double]$IntervalSeconds   # mW * s => mJ
                $accumulated = [double]$pdata.CpuEnergy

                # escape quotes and newlines in process name
                $escName = ($pname -replace '"','""') -replace '\r|\n',' '

                $line2 = "{0},{1},{2:F4},{3:F4},{4:F4}" -f $ts, ('"' + $escName + '"'), $procCpuPct, $procPowerMw, $energyThisSample
                # we want Accumulated as well - append with comma
                $line2 = $line2 + ("," + ([double]$accumulated).ToString("F4", [System.Globalization.CultureInfo]::InvariantCulture))

                $script:CpuProcessesCsvWriter.WriteLine($line2)
            }
        }

        # batched flush behaviour similar to GPU logging
        $script:SamplesSinceLastFlush += 1
        if ($script:SamplesSinceLastFlush -ge $script:SamplesFlushInterval) {
            try {
                if ($null -ne $script:CpuSamplesCsvWriter) { $script:CpuSamplesCsvWriter.Flush() }
                if ($null -ne $script:CpuProcessesCsvWriter) { $script:CpuProcessesCsvWriter.Flush() }

                if ($null -ne $script:SamplesCsvWriter) { $script:SamplesCsvWriter.Flush() }            # GPU samples
                if ($null -ne $script:ProcessesCsvWriter) { $script:ProcessesCsvWriter.Flush() }        # GPU procs
            } catch {}
            $script:SamplesSinceLastFlush = 0
        }
    }
    catch {
        Write-Host "Realtime CPU CSV logging failed: $($_.Exception.Message)"
    }

    
    return $true
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

#region Command Line Interface

function Show-TopList {
    param(
        [int]$Count,
        [bool]$ClearScreen = $true,
        [int]$SampleCount
    )
    
    if ($Count -lt 1) { $Count = 1 }
    if ($Count -gt 50) { $Count = 50 }
    
    $elapsed = $script:Stopwatch.Elapsed.ToString("hh\:mm\:ss")
    
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

    #GPU section
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
    
    # Collect initial CPU data
    for ($i = 0; $i -lt 3; $i++) {
        $null = Update-ProcessEnergyData -IntervalSeconds $IntervalSeconds
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

    # Ensure high-resolution stopwatch is available and warm
    if (-not $script:Stopwatch) {
        $script:Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    } else {
        $script:Stopwatch.Restart()
    }

    if (-not $script:StopwatchFrequency) {
        $script:StopwatchFrequency = [int64][System.Diagnostics.Stopwatch]::Frequency
    }

    # compute GPU sample interval in ticks (SampleIntervalMs is milliseconds)
    $intervalTicksGPU = [int64]([double]$script:SampleIntervalMs * $script:StopwatchFrequency / 1000.0)
    if ($intervalTicksGPU -le 0) { $intervalTicksGPU = 1 }

    # compute CPU measurement interval in ticks (MeasurementInterval is seconds)
    $intervalTicksCPU = [int64]([double]$script:MeasurementInterval * $script:StopwatchFrequency)
    if ($intervalTicksCPU -le 0) { $intervalTicksCPU = 1 }

    # Track last update times (ticks)
    $script:LastSampleTime = $script:Stopwatch.ElapsedTicks
    $LastSampleTimeCpu = $script:Stopwatch.ElapsedTicks

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

            # CPU update: compute actual interval in seconds from ticks
            if ($remainingCPU -le 0) {
                # FIX: $now and $LastSampleTimeCpu are ticks; convert ticks -> seconds
                $deltaTicksCpu = $now - $LastSampleTimeCpu
                if ($deltaTicksCpu -lt 0) { $deltaTicksCpu = 0 }
                $actualIntervalSeconds = [double]$deltaTicksCpu / [double]$script:StopwatchFrequency

                $null = Update-ProcessEnergyData -IntervalSeconds $actualIntervalSeconds

                # update last sample tick for CPU
                $LastSampleTimeCpu = $now

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

Initialize-GPUProcessMonitor -SampleIntervalMs $SampleInterval -wSM $WeightSMP -wMem $WeightMem -wEnc $WeightEnc -wDec $WeightDec -diagPath $DiagnosticsOutputPath
$initialized = Initialize-LibreHardwareMonitor
if (-not $initialized) {
    Write-Host "`nFailed to initialize. Exiting..." -ForegroundColor Red
    exit 1
}

Write-Host "[OK] Hardware monitoring ready" -ForegroundColor Green

Start-CommandLineMode -IntervalSeconds $MeasurementIntervalSeconds

# Cleanup
if ($null -ne $script:Computer) {
    $script:Computer.Close()
}

Write-Host "`nThank you for using Process Power Monitor!`n" -ForegroundColor Cyan

#endregion