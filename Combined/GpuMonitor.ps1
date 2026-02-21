#----------------------------------------------
#region Initialization
#----------------------------------------------

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
    [double]$GpuCurrentPower
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
        $this.GpuCurrentPower = 0.0
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
        try {
            $gpuNameOutput = nvidia-smi --query-gpu=name --format=csv,noheader 2>$null
            $this.GpuName = if ($gpuNameOutput) { $gpuNameOutput.Trim() } else { "unknown" }
        } catch {
            $this.GpuName = "unknown"
            $script:IsGpu = $false
            return 
        }
        Write-Host "Power attribution weights: SM=$wSM, Mem=$wMem, Enc=$wEnc, Dec=$wDec"
        $script:IsGpu = $true

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
        for ($i = 0; $i -lt 100; $i++) {
            try{
                $out = nvidia-smi --query-gpu=power.draw.instant,temperature.gpu,fan.speed --format=csv,noheader,nounits 2>$null
            }
            catch {
                # Write-Log(" $($_.Exception.Message)")
                return
            }
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

    #endregion

    #----------------------------------------------
    #region Utility
    #----------------------------------------------

    [double] ParseUtil([string]$value) {
        if ([string]::IsNullOrWhiteSpace($value) -or $value -eq '-') { return 0.0 }

        $s = $value.Trim()
        $s = $s -replace '[^\d\.\-+]',''

        $d = 0.0
        if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
            return $d
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

    #endregion

    #----------------------------------------------
    #region Core sampling
    #----------------------------------------------

    [array] GetGpuProcesses() {
        if ($null -eq $this.ProcessNameCache) {
            $this.ProcessNameCache = [System.Collections.Generic.Dictionary[int,string]]::new()
            $this.ProcessNameCacheOrder = [System.Collections.Generic.LinkedList[int]]::new()
            if ($this.ProcessNameCacheCapacity -le 0) { $this.ProcessNameCacheCapacity = 1024 }
        }

        $resultList = [System.Collections.Generic.List[object]]::new()
        try{
            $pmonOutput = nvidia-smi pmon -c 1 2>$null
        }
        catch{
            # Write-Host "`n`nNO nvidia-smi!!!!`n" ?
            return $null
        }
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
                SmUtil    = $this.ParseUtil($m.Groups[4].Value)
                MemUtil   = $this.ParseUtil($m.Groups[5].Value)
                EncUtil   = $this.ParseUtil($m.Groups[6].Value)
                DecUtil   = $this.ParseUtil($m.Groups[7].Value)
                Command   = $m.Groups[10].Value.Trim()
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
        $dt = $deltaTicks / $this.StopwatchFrequency

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
        try {
                $combinedOut = nvidia-smi --query-gpu=timestamp,power.draw.instant,utilization.gpu,utilization.memory,utilization.encoder,utilization.decoder,temperature.gpu,fan.speed --format=csv,noheader,nounits 2>$null
        }
        catch {
            return
        }        
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
        $this.GpuCurrentPower = $gpuPower
        $this.GpuEnergyJoules += ($gpuPower * $dt)
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
                        ([double]$perProcessEntry.SMUtil).ToString('F2', $inv) + ',' +
                        ([double]$perProcessEntry.MemUtil).ToString('F2', $inv) + ',' +
                        ([double]$perProcessEntry.EncUtil).ToString('F2', $inv) + ',' +
                        ([double]$perProcessEntry.DecUtil).ToString('F2', $inv) + ',' +
                        ([double]$perProcessEntry.PowerW).ToString('F6', $inv) + ',' +
                        ([double]$perProcessEntry.EnergyJ).ToString('F8', $inv) + ',' +
                        ([double]$perProcessEntry.AccumulatedProcessEnergyJ).ToString('F8', $inv) + ',' +
                        ([double]$perProcessEntry.WeightedUtil).ToString('F4', $inv) + ',' +
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

    #endregion

    #----------------------------------------------
    #region Interface
    #----------------------------------------------

    [void] Report([int]$SampleCount) {
        $Duration = $this.Stopwatch.Elapsed.TotalSeconds
        $safeDuration = if ($Duration -le 0 -or [double]::IsNaN($Duration)) { 1.0 } else { [double]$Duration }
        $avgGpuPower = if ($safeDuration -gt 0) { $this.GpuEnergyJoules / $safeDuration } else { 0.0 }

        $RuntimeStr = "{0:hh\:mm\:ss}" -f $this.Stopwatch.Elapsed
        # System stats
        Write-Host "GPU Diagnostics:" -ForegroundColor $script:Colors.Info
        Write-Host ("  Runtime:              {0:hh\:mm\:ss}" -f $RuntimeStr) -ForegroundColor $script:Colors.Value
        Write-Host ("  Measurements:         {0}" -f $SampleCount) -ForegroundColor $script:Colors.Value
        Write-Host ("  Measurement Interval: {0}s" -f $this.SampleIntervalMs) -ForegroundColor $script:Colors.Value
        Write-Host ("  Tracked Processes:    {0}" -f $this.ProcessEnergyJoules.Count) -ForegroundColor $script:Colors.Value
        Write-Host ("  Idle GPU Power:       {0:N2}W [Min: {1:N2}W  Max:{2:N2}W] with {3} Processes" -f $this.GpuIdlePower, $this.GpuIdlePowerMin, $this.GpuIdlePowerMax, $this.IdleProcessCount) -ForegroundColor Yellow
        Write-Host ("  Idle GPU Temperature: {0:F1}C" -f $this.GpuIdleTemperatureC) -ForegroundColor Yellow
        Write-Host ("  Idle GPU Fan Util.:   {0:F1}%" -f $this.GpuIdleFanPercent) -ForegroundColor Yellow
        Write-Host ("  GPU Power Now:        {0:F2}W" -f $this.GpuCurrentPower) -ForegroundColor $script:Colors.Value
        Write-Host ("  Accumulative Average: {0:F2}W" -f $avgGpuPower) -ForegroundColor $script:Colors.Value
        Write-Host ("  Total GPU Energy:     {0:F2}J" -f $this.GpuEnergyJoules) -ForegroundColor Yellow
        
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
                foreach ($pr in $procs) { $pidToName[[int]$pr.Id] = $pr.ProcessName }
            } catch {}
        }

        # Fallback: if Get-Process didn't return a name (process ended), try to obtain the name
        if ($allPids.Count -gt 0) {
            foreach ($procId in $allPids) {
                if (-not $pidToName.ContainsKey($procId)) {
                    $foundName = $null
                    foreach ($s in $this.Samples) {
                        if ($s.PerProcess -and $s.PerProcess.ContainsKey([string]$procId)) {
                            $pp = $s.PerProcess[[string]$procId]
                            if ($pp -and $pp.ProcessName) {
                                $foundName = [string]$pp.ProcessName
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
        foreach ($e in $entriesList) { $sumTotalProcessesEnergy += [double]$e.Value }

        # Diagnostic compare total vs sum of processes
        if ($sumTotalProcessesEnergy -gt 0) {
            $diff = [math]::Abs([double]$this.GpuEnergyJoules - $sumTotalProcessesEnergy)
            $diffPct = if ($this.GpuEnergyJoules -gt 0) { 100.0 * $diff / $this.GpuEnergyJoules } else { 0.0 }
            Write-Host("  Measured GPU total {0:F2}J vs Summed Process Attributed {1:F2}J -> Diff: {2:F2}J ({3:F1}%)" -f $this.GpuEnergyJoules, $sumTotalProcessesEnergy, $diff, $diffPct) -ForegroundColor Yellow
        }
        else {
            Write-Host ""
        }
        Write-Host ""

        Write-Host "Top GPU Processes:" -ForegroundColor $script:Colors.Highlight
        Write-Host ("{0,-15} {1,-30} {2,-25:F2} {3,-20:F2} {4:F1}%" -f "PID/s", "Process Name", "Accumulated Energy (J)", "Average Power (W)", "All Time Energy Contribution (%)") -ForegroundColor $script:Colors.Header
        Write-Host ("-" * 75) -ForegroundColor $script:Colors.Header

        # Aggregate by process name and print
        $agg = @{}
        foreach ($e in $entriesList) {
            $pidNum = [int]$e.Key
            $pname = if ($pidToName.ContainsKey($pidNum)) { $pidToName[$pidNum] } else { "[exited] PID $pidNum" }
            if (-not $agg.ContainsKey($pname)) { $agg[$pname] = @{ Energy = 0.0; Pids = New-Object 'System.Collections.Generic.List[int]' } }
            $agg[$pname].Energy += [double]$e.Value
            [void]$agg[$pname].Pids.Add($pidNum)
        }

        $agg.GetEnumerator() | Sort-Object { $_.Value.Energy } -Descending | ForEach-Object {
            $name = $_.Key
            $energy = $_.Value.Energy
            $avgPowerProc = if ($safeDuration -gt 0) { $energy / $safeDuration } else { 0.0 }
            $pct = if ($sumTotalProcessesEnergy -gt 0) { 100.0 * $energy / $sumTotalProcessesEnergy } else { 0.0 }
            $pidList = ($_.Value.Pids -join ',')
            Write-Host ("{0,-15} {1,-30} {2,-25:F2} {3,-20:F2} {4:F1}%" -f $pidList, $name, $energy, $avgPowerProc, $pct) -ForegroundColor Yellow
        }
    }
    #endregion
}