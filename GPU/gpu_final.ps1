<#
Usage example:
For general monitoring of all GPU processes:
.\gpu_final.ps1
For specific process by name or procId (e.g., "python" or "1234"):
.\gpu_final.ps1 -Process "python"
Other flags of interest:
    -Duration: how long to monitor in seconds (default 60)
    -SampleInterval: how often to sample in ms (default 100)
    -WeightSM, WeightMem, WeightEnc, WeightDec: weights for the attribution formula (defaults: 1.0, 0.5, 0.25, 0.15)
    -DiagnosticsOutput: base path for writing diagnostics CSV files (default "gpu_diagnostics")
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

        # Get GPU name
        $gpuNameOutput = nvidia-smi --query-gpu=name --format=csv,noheader 2>$null
        $this.GpuName = $gpuNameOutput.Trim()
        Write-Log "Monitoring GPU: $($this.GpuName)"
        Write-Host "Power attribution weights: SM=$wSM, Mem=$wMem, Enc=$wEnc, Dec=$wDec"

        # Measure idle metrics
        Write-Log "Measuring idle power (please ensure GPU is idle; close heavy apps)."
        $this.MeasureIdleMetrics()
        Write-Log (("Idle GPU power measured: {0:N2}W [Min: {1:N2}W  Max:{2:N2}W] with {3} Processes; Temp: {4:F1}C  Fan: {5:F1}%" -f `
            $this.GpuIdlePower, $this.GpuIdlePowerMin, $this.GpuIdlePowerMax, $this.IdleProcessCount, $this.GpuIdleTemperatureC, $this.GpuIdleFanPercent))
    }

    [void] MeasureIdleMetrics() {
        # Take multiple samples to get a stable idle measurement using a single batched nvidia-smi call
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

    # Faster GetGpuProcesses(): compiled regex, List<>, HashSet for PIDs, single Get-Process
    [array] GetGpuProcesses() {
        $processList = @()

        $pmonOutput = nvidia-smi pmon -c 1 2>$null
        if (-not $pmonOutput) { return $processList }

        $parsed = @()
        foreach ($line in ($pmonOutput -split "`n")) {
            if ($line -match '^\s*#' -or $line -match '^-+' -or [string]::IsNullOrWhiteSpace($line)) { continue }
            if ($line -match '^\s*(\d+)\s+(\d+)\s+([A-Z+]+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(.+?)\s*$') {
                $parsed += [PSCustomObject]@{
                    ProcessId = [int]$Matches[2]
                    SmUtil    = $this.ParseUtil($Matches[4])
                    MemUtil   = $this.ParseUtil($Matches[5])
                    EncUtil   = $this.ParseUtil($Matches[6])
                    DecUtil   = $this.ParseUtil($Matches[7])
                    Command   = $Matches[10].Trim()
                }
            }
        }

        if ($parsed.Count -eq 0) { return $processList }

        # Resolve Windows process names ONCE (same as working version)
        $procNameMap = @{}
        $pids = $parsed | Select-Object -ExpandProperty ProcessId -Unique
        try {
            $procs = Get-Process -Id $pids -ErrorAction SilentlyContinue
            foreach ($pr in $procs) {
                $procNameMap[[int]$pr.Id] = $pr
            }
        }
        catch {}

        foreach ($entry in $parsed) {
            $procObj = $null
            if ($procNameMap.ContainsKey($entry.ProcessId)) {
                $procObj = $procNameMap[$entry.ProcessId]
            }

            $processName = if ($procObj) {
                $procObj.ProcessName
            }
            else {
                $entry.Command
            }

            $processList += [PSCustomObject]@{
                ProcessId    = $entry.ProcessId
                ProcessName  = [string]$processName
                SmUtil       = $entry.SmUtil
                MemUtil      = $entry.MemUtil
                EncUtil      = $entry.EncUtil
                DecUtil      = $entry.DecUtil
                WeightedUtil = 0.0
            }
        }

        return $processList
    }

    [double] ParseUtil([string]$value) {
        if ([string]::IsNullOrWhiteSpace($value) -or $value -eq '-') { return 0.0 }
        if ($value -match '([-+]?[0-9]*\.?[0-9]+)') {
            try {
                return [double]$matches[1]
            }
            catch {}
        }
        return 0.0
    }

    [void] Sample() {
        # Cache frequently used instance fields locally to avoid repeated property resolution
        $weightSmLocal      = $this.WeightSm
        $weightMemoryLocal  = $this.WeightMemory
        $weightEncoderLocal = $this.WeightEncoder
        $weightDecoderLocal = $this.WeightDecoder
        $gpuIdleProcessMap  = $this.GpuIdleProcesses
        $processEnergyMap   = $this.ProcessEnergyJoules

        # High-resolution monotonic timing (fast path)
        $currentTicks = $this.Stopwatch.ElapsedTicks
        $deltaTicks   = $currentTicks - $this.LastSampleTime
        $this.LastSampleTime = $currentTicks
        if ($deltaTicks -le 0) { return }

        # Convert to seconds (double, sub-ms precision)
        $dt = $deltaTicks / $this.StopwatchFrequency

        # Batched nvidia-smi for GPU-wide metrics
        $timestamp = "-"
        $gpuPower = 0.0
        $gpuSmUtil = 0.0
        $gpuMemUtil = 0.0
        $gpuEncUtil = 0.0
        $gpuDecUtil = 0.0
        $gpuTemperature = 0.0
        $gpuFanUtil = 0.0
        $combinedOut = nvidia-smi --query-gpu=timestamp,power.draw.instant,utilization.gpu,utilization.memory,utilization.encoder,utilization.decoder,temperature.gpu,fan.speed --format=csv,noheader,nounits 2>$null
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
            }
            catch {}
        }

        # Get running processes with their utilizations (fast)
        $processEntries = $this.GetGpuProcesses()
        $currentProcessCount = $processEntries.Count

        # Compute weighted GPU total
        $gpuWeightedTotal = ($weightSmLocal * $gpuSmUtil) +
                            ($weightMemoryLocal * $gpuMemUtil) +
                            ($weightEncoderLocal * $gpuEncUtil) +
                            ($weightDecoderLocal * $gpuDecUtil)

        # Per-process weighted totals
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
            $isIdleProcessFlag = $gpuIdleProcessMap.ContainsKey($processIdStringForLookup)
            if ($weightedValue -le 0 -and $isIdleProcessFlag) {
                $inactiveIdleProcessesCount++
            }
        }

        # Attribution math (unchanged)
        $gpuActivePower = if (($this.GpuIdlePower) -lt $gpuPower) { $gpuPower - ($this.GpuIdlePower) } else { 0 }
        $gpuTotalProcessUtil = if ($gpuWeightedTotal -gt 0) { $processWeightTotal / $gpuWeightedTotal } else { 0 }
        $gpuTotalActiveProcessPower = $gpuTotalProcessUtil * $gpuActivePower
        $gpuExcessPower = $gpuActivePower - $gpuTotalActiveProcessPower

        $activeProcessCount = if ($inactiveIdleProcessesCount -eq $currentProcessCount) { $currentProcessCount } else { $currentProcessCount - $inactiveIdleProcessesCount }
        $P_idle_pwr = if ($this.IdleProcessCount -gt 0) { $this.GpuIdlePower / $this.IdleProcessCount } else { 0 }
        $P_residual_per_proc = if ($activeProcessCount -gt 0) { $gpuExcessPower / $activeProcessCount } else { 0 }

        # Build per-process attribution objects quickly
        $sampleAttributedPower = 0.0
        $samplePerProcess = @{}
        foreach ($processEntry in $processEntries) {
            $processIdString = $processEntry.ProcessId.ToString()
            $weightedForThisProcess = $processEntry.WeightedUtil

            if ($processWeightTotal -le 0) { $fraction = 0 } else { $fraction = $weightedForThisProcess / $processWeightTotal }

            if ($gpuIdleProcessMap.ContainsKey($processIdString)) {
                if ($weightedForThisProcess -le 0) {
                    $powerValue = if ($inactiveIdleProcessesCount -eq $currentProcessCount) { $P_idle_pwr + $P_residual_per_proc } else { $P_idle_pwr }
                }
                else {
                    $powerValue = $P_idle_pwr + ($fraction * $gpuTotalActiveProcessPower) + $P_residual_per_proc
                }
            }
            else {
                if ($weightedForThisProcess -le 0) {
                    $powerValue = $P_residual_per_proc
                }
                else {
                    $powerValue = ($fraction * $gpuTotalActiveProcessPower) + $P_residual_per_proc
                }
            }

            $energyJ = $powerValue * $dt

            # safe and faster update of ProcessEnergyJoules
            if ($processEnergyMap.ContainsKey($processIdString)) {
                $existingEnergyValue = [double]$processEnergyMap[$processIdString]
            } else {
                $existingEnergyValue = 0.0
            }
            $processEnergyMap[$processIdString] = $existingEnergyValue + $energyJ

            $sampleAttributedPower += $powerValue

            $samplePerProcess[$processIdString] = [PSCustomObject]@{
                PID          = $processIdString
                ProcessName  = [string]$processEntry.ProcessName
                SMUtil       = $processEntry.SmUtil
                MemUtil      = $processEntry.MemUtil
                EncUtil      = $processEntry.EncUtil
                DecUtil      = $processEntry.DecUtil
                PowerW       = [math]::Round($powerValue, 4)
                EnergyJ      = [math]::Round($energyJ, 6)
                WeightedUtil = $weightedForThisProcess
                IsIdle       = [bool]$gpuIdleProcessMap.ContainsKey($processIdString)
            }
        }

        # Record sample
        $this.GpuEnergyJoules += ($gpuPower * $dt)
        $sampleResidualPower = $gpuPower - $sampleAttributedPower
        [void]$this.Samples.Add([PSCustomObject]@{
                Timestamp         = $timestamp
                PowerW            = $gpuPower
                ActivePowerW      = $gpuActivePower
                ExcessPowerW      = $gpuExcessPower
                GpuSmUtil         = $gpuSmUtil
                GpuMemUtil        = $gpuMemUtil
                GpuEncUtil        = $gpuEncUtil
                GpuDecUtil        = $gpuDecUtil
                GpuWeightedTotal  = $gpuWeightedTotal
                ProcessWeightTotal= $processWeightTotal
                ProcessCount      = $currentProcessCount
                AttributedPowerW  = [math]::Round($sampleAttributedPower, 4)
                ResidualPowerW    = [math]::Round($sampleResidualPower, 4)
                TemperatureC      = $gpuTemperature
                FanPercent        = $gpuFanUtil
                PerProcess        = $samplePerProcess
            })

        # --- Real-time diagnostics logging (append-only, per-sample) -------------
        try {
            # mirror Report() behavior for paths
            $timestampForFile = $this.TimeStampLogging.ToString('yyyyMMdd_HHmmss')
            $outSpec = $this.DiagnosticsOutputPath

            if (-not $outSpec) {
                $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
                $outSpec = Join-Path $scriptDir 'gpu_diagnostics'
            }

            if ($outSpec -match '\.csv$') {
                $csvPath = $outSpec
                $csvPathProcesses = $outSpec -replace '\.csv$','_processes.csv'
                $diagDir = Split-Path -Parent $csvPath
                if ($diagDir -and -not (Test-Path $diagDir)) { New-Item -ItemType Directory -Path $diagDir -Force | Out-Null }
            }
            else {
                $diagDir = $outSpec
                if (-not (Test-Path $diagDir)) { New-Item -ItemType Directory -Path $diagDir -Force | Out-Null }
                $csvPath = Join-Path $diagDir ("samples_$timestampForFile.csv")
                $csvPathProcesses = Join-Path $diagDir ("processes_$timestampForFile.csv")
            }

            # prepare invariant culture for number formatting
            $inv = [System.Globalization.CultureInfo]::InvariantCulture

            # prepare sample-line values (escape quotes in timestamp)
            $tsEsc = ($timestamp -replace '"','""')
            $pw   = ([double]$gpuPower).ToString('F4', $inv)
            $act  = ([double]$gpuActivePower).ToString('F4', $inv)
            $ex   = ([double]$gpuExcessPower).ToString('F4', $inv)
            $sm   = ([double]$gpuSmUtil).ToString('F2', $inv)
            $mem  = ([double]$gpuMemUtil).ToString('F2', $inv)
            $enc  = ([double]$gpuEncUtil).ToString('F2', $inv)
            $dec  = ([double]$gpuDecUtil).ToString('F2', $inv)
            $gwt  = ([double]$gpuWeightedTotal).ToString('F2', $inv)
            $pwt  = ([double]$processWeightTotal).ToString('F2', $inv)
            $pc   = [int]$currentProcessCount
            $apow = ([double]$sampleAttributedPower).ToString('F4', $inv)
            $rpow = ([double]$sampleResidualPower).ToString('F4', $inv)
            $tmpC = ([double]$gpuTemperature).ToString('F1', $inv)
            $fanP = ([double]$gpuFanUtil).ToString('F1', $inv)

            # write sample CSV (append, create header if needed)
            $writeHeaderSample = -not (Test-Path $csvPath)
            $sw = [System.IO.StreamWriter]::new($csvPath, $true, [System.Text.Encoding]::UTF8)
            try {
                if ($writeHeaderSample) {
                    $sw.WriteLine('Timestamp,PowerW,ActivePowerW,ExcessPowerW,GpuSMUtil,GpuMemUtil,GpuEncUtil,GpuDecUtil,GpuWeightTotal,ProcessWeightTotal,ProcessCount,AttributedPowerW,ResidualPowerW,TemperatureC,FanPercent')
                }
                $line = '"' + $tsEsc + '",' +
                        $pw  + ',' +
                        $act + ',' +
                        $ex  + ',' +
                        $sm  + ',' +
                        $mem + ',' +
                        $enc + ',' +
                        $dec + ',' +
                        $gwt + ',' +
                        $pwt + ',' +
                        $pc  + ',' +
                        $apow + ',' +
                        $rpow + ',' +
                        $tmpC + ',' +
                        $fanP
                $sw.WriteLine($line)
            } finally { $sw.Close() }

            # write per-process CSV (append, create header if needed)
            $writeHeaderProc = -not (Test-Path $csvPathProcesses)
            $sw2 = [System.IO.StreamWriter]::new($csvPathProcesses, $true, [System.Text.Encoding]::UTF8)
            try {
                if ($writeHeaderProc) {
                    $sw2.WriteLine('Timestamp,PID,ProcessName,SMUtil,MemUtil,EncUtil,DecUtil,PowerW,EnergyJ,WeightedUtil,IsIdle')
                }
                foreach ($perProcessEntry in $samplePerProcess.Values) {
                    $p_tsEsc = ($timestamp -replace '"','""')
                    $procId = $perProcessEntry.PID
                    $pnameEsc = ($perProcessEntry.ProcessName -replace '"','""') -replace '\r|\n',' '
                    $smv = ([double]$perProcessEntry.SMUtil).ToString('F2', $inv)
                    $memv = ([double]$perProcessEntry.MemUtil).ToString('F2', $inv)
                    $encv = ([double]$perProcessEntry.EncUtil).ToString('F2', $inv)
                    $decv = ([double]$perProcessEntry.DecUtil).ToString('F2', $inv)
                    $pwr = ([double]$perProcessEntry.PowerW).ToString('F4', $inv)
                    $enj = ([double]$perProcessEntry.EnergyJ).ToString('F6', $inv)
                    $wut = ([double]$perProcessEntry.WeightedUtil).ToString('F4', $inv)
                    $idl = [bool]$perProcessEntry.IsIdle

                    $line2 = '"' + $p_tsEsc + '",' +
                             $procId + ',"' + $pnameEsc + '",' +
                             $smv + ',' +
                             $memv + ',' +
                             $encv + ',' +
                             $decv + ',' +
                             $pwr + ',' +
                             $enj + ',' +
                             $wut + ',' +
                             $idl
                    $sw2.WriteLine($line2)
                }
            } finally { $sw2.Close() }
        }
        catch {
            # don't throw — real-time logging should not break sampling
            Write-Log "Realtime CSV logging failed: $($_.Exception.Message)"
        }
        # --- end real-time logging ------------------------------------------------
    }

    # Run method: Duration is optional (null = run forever)
    [void] Run([Nullable[int]]$Duration) {
        # Prompt user to start workload
        Write-Host ""; Write-Log "READY: Start the workload now (e.g., compute or memory workload)."
        Write-Host "Press ENTER when the workload is running and you want to begin sampling..."
        [void][System.Console]::ReadLine()

        # Ensure high-resolution stopwatch is available and warm
        if (-not $this.Stopwatch) {
            $this.Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        }
        else {
            $this.Stopwatch.Restart()
        }

        # Cache frequency for faster conversions in hot path
        if (-not $this.StopwatchFrequency) {
            $this.StopwatchFrequency = [int64][System.Diagnostics.Stopwatch]::Frequency
        }

        # Initialize last-sample ticks for the Sample() method
        $this.LastSampleTime = $this.Stopwatch.ElapsedTicks

        # If Duration provided, compute end ticks; if not, run forever
        $runForever = $false
        $endTicks = 0
        if ($null -eq $Duration) {
            $runForever = $true
            Write-Log "Monitoring GPU indefinitely (no duration specified). Use Ctrl+C to stop."
        }
        else {
            # Convert seconds -> ticks (int64)
            $endTicks = [int64]($Duration * $this.StopwatchFrequency)
            Write-Log "Monitoring GPU for $Duration seconds..."
            # Optional: log the expected end time in wall-clock terms
            $endWallClock = (Get-Date).AddSeconds([double]$Duration)
            Write-Log "End time (wall clock): $endWallClock"
        }

        $sampleCount = 0
        try {
            if ($runForever) {
                while ($true) {
                    $this.Sample()
                    $sampleCount++
                    if ($this.SampleIntervalMs -gt 0) {
                        Start-Sleep -Milliseconds $this.SampleIntervalMs
                    }
                }
            }
            else {
                # Stop when elapsed ticks >= endTicks
                while ($this.Stopwatch.ElapsedTicks -lt $endTicks) {
                    $this.Sample()
                    $sampleCount++
                    if ($this.SampleIntervalMs -gt 0) {
                        Start-Sleep -Milliseconds $this.SampleIntervalMs
                    }
                }
            }
        }
        catch [System.Exception] {
            # Allow Ctrl+C or other termination to bubble, but report gracefully
            Write-Log "Monitoring stopped: $($_.Exception.Message)"
        }
        finally {
            $Duration = $this.Stopwatch.Elapsed.TotalSeconds
            Write-Log "Collected $sampleCount samples over $Duration seconds."
            $this.Report($Duration)
        }
    }

    [void] Report([double]$Duration) {
        Write-Host "`n==== GPU POWER SUMMARY ===="

        # Safe duration for averages
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
                foreach ($pr in $procs) { $pidToName[[int]$pr.Id] = $pr.ProcessName }
            } catch { }
        }

        # Fallback: if Get-Process didn't return a name (process ended), try to obtain the name
        # from recorded samples' PerProcess entries (first seen name wins).
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
            Write-Log ("DIAGNOSTIC: Measured total {0:F2} J vs summed attributed {1:F2} J -> Diff: {2:F2} J ({3:F1}%)" -f $this.GpuEnergyJoules, $sumTotalProcessesEnergy, $diff, $diffPct)
        }

        # Aggregate by process name and print (same logic as before)
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
            Write-Host ("{0,-30} Accumulative: {1,8:F2} J  |  Average: {2,6:F2} W  ({3,5:F1}%)  PIDs: {4}" -f $name, $energy, $avgPowerProc, $pct, $pidList)
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
    exit 1
}