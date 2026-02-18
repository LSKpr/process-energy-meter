<#
Usage example:
.\gpu_complete.ps1 -Process "myapp" -Duration 60 -SampleInterval 100 -WeightSm 1.0 -WeightMem 0.5 -WeightEnc 0.25 -WeightDec 0.15 -AutoScale -Diagnostics
#>

param(
    [string]$Process = $null,
    [ValidateRange(1,86400)][int]$Duration = 60,
    [ValidateRange(10,60000)][int]$SampleInterval = 100,
    [double]$WeightSM = 1.0,
    [double]$WeightMem = 0.5,
    [double]$WeightEnc = 0.25,
    [double]$WeightDec = 0.15,
    [switch]$AutoScale,
    [switch]$Diagnostics,
    [string]$DiagnosticsOutput = $null
)

function Write-Log($msg) {
    $out = nvidia-smi --query-gpu=timestamp --format=csv,noheader 2>$null
    if ($out -and $out.Trim() -ne '[N/A]') {
        try {
            $t = [datetime]::ParseExact(
                $out.Trim(),
                'yyyy/MM/dd HH:mm:ss.fff',
                [System.Globalization.CultureInfo]::InvariantCulture
            )
            $tStr = $t.ToString('yyyy/MM/dd HH:mm:ss.fff')
            Write-Host "[$tStr] $msg"
        }
        catch {}
    }
    else {
        $t = Get-Date
        $tStr = $t.ToString('yyyy/MM/dd HH:mm:ss.fff')
        Write-Host "[$tStr] $msg"
    }
}

class GPUProcessMonitor {
    [string]$TargetProcess
    [int]$SampleIntervalMs
    [double]$WeightSm      # SM weight
    [double]$WeightMemory  # Memory weight
    [double]$WeightEncoder # Encoder weight
    [double]$WeightDecoder # Decoder weight

    [System.Collections.Generic.List[object]]$Samples
    [hashtable]$ProcessEnergyJoules
    [datetime]$LastSampleTime

    [string]$GpuName
    [double]$GpuEnergyJoules
    [double]$GpuIdlePower
    [double]$GpuIdleMemoryUsedMB
    [double]$GpuIdleFanPercent
    [double]$GpuIdleTemperatureC
    [double]$GpuTotalMemoryMB
    [hashtable]$GpuIdleProcesses
    [int]$IdleProcessCount
    [bool]$IsDriverInWDDM
    [bool]$AutoScaleEnabled
    [bool]$DiagnosticsEnabled
    [string]$DiagnosticsOutputPath

    GPUProcessMonitor([string]$targetProcess, [int]$SampleIntervalMs, [double]$wSM, [double]$wMem, [double]$wEnc, [double]$wDec, [bool]$autoScale, [bool]$diagEnabled, [string]$diagPath) {
        $this.TargetProcess = $targetProcess
        $this.SampleIntervalMs = $SampleIntervalMs
        $this.WeightSm = $wSM
        $this.WeightMemory = $wMem
        $this.WeightEncoder = $wEnc
        $this.WeightDecoder = $wDec
        
        $this.Samples = [System.Collections.Generic.List[object]]::new()
        $this.ProcessEnergyJoules = @{}
        $this.GpuEnergyJoules = 0.0
        $this.GpuIdleProcesses = @{}
        $this.LastSampleTime = Get-Date

        $this.GpuIdleMemoryUsedMB = 0.0
        $this.IsDriverInWDDM = $this.IsWDDM()
        $this.AutoScaleEnabled = $autoScale
        $this.DiagnosticsEnabled = $diagEnabled
        $this.DiagnosticsOutputPath = $diagPath

        # Get GPU name and total memory
        $this.GpuName = $this.GetGpuName()
        $this.GpuTotalMemoryMB = $this.GetMemoryTotal()
        Write-Log "Monitoring GPU: $($this.GpuName) with total memory: $([math]::Round($this.GpuTotalMemoryMB,1)) MiB"
        Write-Host "Power attribution weights: SM=$wSM, Mem=$wMem, Enc=$wEnc, Dec=$wDec"

        # Measure idle metrics
        $this.MeasureIdleMetrics()
    }

    [void] MeasureIdleMetrics() {
        Write-Log "Measuring idle power (please ensure GPU is idle; close heavy apps)."
        # Take multiple samples to get a stable idle measurement using a single batched nvidia-smi call
        $idleTemperatureSamples = @()
        $idleFanUtilSamples = @()
        $idleMemorySamples = @()
        $idlePowerSamples = @()
        $idlePowerMin = 0
        $idlePowerMax = 0
        for ($i = 0; $i -lt 10; $i++) {
            $out = nvidia-smi --query-gpu=power.draw.instant,memory.used,temperature.gpu,fan.speed --format=csv,noheader,nounits 2>$null
            if ($out -and -not [string]::IsNullOrWhiteSpace($out) -and $out.Trim() -ne '[N/A]') {
                $parts = $out -split ',\s*'
                if ($parts.Count -ge 4) {
                    try {
                        $idlePowerSamples += [double]$parts[0].Trim()
                        $idleMemorySamples += [double]$parts[1].Trim()
                        $idleTemperatureSamples += [double]$parts[2].Trim()
                        $idleFanUtilSamples += [double]$parts[3].Trim()
                    }
                    catch {}
                }
            }
            Start-Sleep -Milliseconds $this.SampleIntervalMs
        }

        if ($idleTemperatureSamples.Count -gt 0) { $this.GpuIdleTemperatureC = ($idleTemperatureSamples | Measure-Object -Average).Average }
        if ($idleFanUtilSamples.Count -gt 0) { $this.GpuIdleFanPercent = ($idleFanUtilSamples | Measure-Object -Average).Average }
        if ($idleMemorySamples.Count -gt 0) { $this.GpuIdleMemoryUsedMB = ($idleMemorySamples | Measure-Object -Average).Average }
        if ($idlePowerSamples.Count -gt 0) {
            $measuredIdlePower = ($idlePowerSamples | Measure-Object -Average).Average
            $idlePowerMin = $idlePowerSamples | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
            $idlePowerMax = $idlePowerSamples | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
            if ($measuredIdlePower -gt 50.0) {
                $this.GpuIdlePower = 15.0 # Fallback for modern GPUs
                Write-Host "WARNING: Measured idle power ($([double]$measuredIdlePower).ToString('F1') W) is suspiciously high." -ForegroundColor Red
                Write-Host "       This usually means that the GPU is already under load."
                Write-Host "       Defaulting Idle Power to 15.0 W to ensure correct attribution." -ForegroundColor Yellow
            }
            else { $this.GpuIdlePower = $measuredIdlePower }
        }

        ##############
        # Record which processes were running during idle measurement
        $idleProcesses = $this.GetGpuProcesses()
        $this.IdleProcessCount = $idleProcesses.Count
        foreach ($proc in $idleProcesses) {
            $this.GpuIdleProcesses[[string]$proc.ProcessId] = $true
        }

        Write-Log (("Idle GPU power measured ({3} samples): {0:N2}W [Min: {1:N2}W  Max:{2:N2}W] with {4} Processes; Temp: {5:F1}C  Fan: {6:F1}%  IdleMem: {7:F1}MiB" -f `
            $this.GpuIdlePower, $idlePowerMin, $idlePowerMax, $idlePowerSamples.Count, $this.IdleProcessCount, $this.GpuIdleTemperatureC, $this.GpuIdleFanPercent, $this.GpuIdleMemoryUsedMB))
    }

    [array] GetGpuProcesses() {
        $processes = @()

        # Use nvidia-smi pmon to get per-process utilization (parse first)
        $pmonOutput = nvidia-smi pmon -c 1 2>$null
        if (-not $pmonOutput) { return $processes }

        $parsed = @()
        foreach ($line in $pmonOutput -split "`n") {
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

        if ($parsed.Count -eq 0) { return $processes }

        # Build PID list and query memory for all compute-apps once
        $pids = $parsed | Select-Object -ExpandProperty ProcessId -Unique
        $memMap = @{}
        try {
            $computeOutput = nvidia-smi --query-compute-apps=pid,used_memory --format=csv,noheader,nounits 2>$null
            if ($computeOutput) {
                foreach ($line in $computeOutput -split "`n") {
                    if ([string]::IsNullOrWhiteSpace($line)) { continue }
                    $parts = $line -split ',\s*'
                    if ($parts.Count -ge 2) {
                        $procId = 0
                        try { $procId = [int]$parts[0].Trim() } catch { continue }
                        $mem = $parts[1].Trim()
                        if ($mem -ne '[N/A]' -and $mem -match '^\d+') { $memMap[[int]$procId] = [double]$mem }
                    }
                }
            }
        }
        catch {}

        # Query Windows processes once to get names and working set memory
        $procNameMap = @{}
        try {
            $procs = Get-Process -Id $pids -ErrorAction SilentlyContinue
            foreach ($pr in $procs) { $procNameMap[[int]$pr.Id] = $pr }
        }
        catch {}

        foreach ($entry in $parsed) {
            $procId = [int]$entry.ProcessId
            $procObj = $null
            if ($procNameMap.ContainsKey($procId)) { $procObj = $procNameMap[$procId] }
            $processName = if ($procObj) { $procObj.ProcessName } else { $entry.Command }

            $memMB = 0.0
            if ($memMap.ContainsKey($procId)) { $memMB = $memMap[$procId] }
            elseif ($procObj) { $memMB = [math]::Round($procObj.WorkingSet / 1MB, 2) }

            $processes += [PSCustomObject]@{
                ProcessId    = $procId
                ProcessName  = $processName
                UsedMemoryMB = $memMB
                SmUtil       = $entry.SmUtil
                MemUtil      = $entry.MemUtil
                EncUtil      = $entry.EncUtil
                DecUtil      = $entry.DecUtil
            }
        }

        return $processes
    }

    [bool] IsProcessRunning() {
        $processes = $this.GetGpuProcesses()
        foreach ($proc in $processes) {
            if ($proc.ProcessName -like "*$($this.TargetProcess)*") {
                return $true
            }
        }
        return $false
    }

    [bool] IsWDDM() {
        $gpuDriverModelOutput = nvidia-smi --query-gpu=driver_model.current --format=csv,noheader 2>$null
        if (-not $gpuDriverModelOutput) { return $false }
        return ($gpuDriverModelOutput.Trim().ToUpperInvariant() -eq 'WDDM')
    }

    [string] GetGpuName() {
        $gpuNameOutput = nvidia-smi --query-gpu=name --format=csv,noheader 2>$null
        if (-not $gpuNameOutput -or $LASTEXITCODE -ne 0) {
            Write-Error "nvidia-smi not found or failed to execute"
            exit 1
        }
        return $gpuNameOutput.Trim()
    }

    [double] GetGpuPower() {
        $out = nvidia-smi --query-gpu=power.draw.instant --format=csv,noheader,nounits 2>$null
        if ($out -and $out.Trim() -ne '[N/A]') {
            try { return [double]$out.Trim() } catch {}
        }
        return 0.0
    }

    [double] GetMemoryTotal() {
        $out = nvidia-smi --query-gpu=memory.total --format=csv,noheader,nounits 2>$null
        if ($out -and $out.Trim() -ne '[N/A]') {
            try { return [double]$out.Trim() } catch {}
        }
        return 0.0
    }

    [double] GetMemoryUsed() {
        $out = nvidia-smi --query-gpu=memory.used --format=csv,noheader,nounits 2>$null
        if ($out -and $out.Trim() -ne '[N/A]') {
            try { return [double]$out.Trim() } catch {}
        }
        return 0.0
    }

    [double] GetTemperature() {
        $out = nvidia-smi --query-gpu=temperature.gpu --format=csv,noheader,nounits 2>$null
        if ($out -and $out.Trim() -ne '[N/A]') {
            try { return [double]$out.Trim() } catch {}
        }
        return 0.0
    }

    [double] GetFanUtil() {
        $out = nvidia-smi --query-gpu=fan.speed --format=csv,noheader,nounits 2>$null
        if ($out -and $out.Trim() -ne '[N/A]') {
            try { return [double]$out.Trim() } catch {}
        }
        return 0.0
    }

    [datetime] GetGpuTime() {
        $out = nvidia-smi --query-gpu=timestamp --format=csv,noheader 2>$null
        if ($out -and $out.Trim() -ne '[N/A]') {
            try {
                return [datetime]::ParseExact(
                    $out.Trim(),
                    'yyyy/MM/dd HH:mm:ss.fff',
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
            }
            catch {}
        }
        return [datetime](Get-Date)
    }

    # GetProcessMemory is no longer used; memory is obtained in GetGpuProcesses() in a batched fashion.

    [double] ParseUtil([string]$value) {
        if ([string]::IsNullOrWhiteSpace($value) -or $value -eq '-') { return 0.0 }
        if ($value -match '([-+]?[0-9]*\.?[0-9]+)') { try { return [double]$matches[1] } catch { } }
        return 0.0
    }

    [void] Sample() {
        # Calculate ACTUAL time elapsed since last sample (ms precision)
        $now = Get-Date
        $dtMilliseconds = ($now - $this.LastSampleTime).TotalMilliseconds
        $dt = $dtMilliseconds / 1000.0
        $this.LastSampleTime = $now
        if ($dt -le 0) { return }
        
        # Get current GPU power draw and overall GPU utilization in one batched nvidia-smi call
        $gpuSmUtil = 0.0; $gpuMemUtil = 0.0; $gpuEncUtil = 0.0; $gpuDecUtil = 0.0
        $gpuPower = 0.0
        $combinedOut = nvidia-smi --query-gpu=power.draw.instant,utilization.gpu,utilization.memory,utilization.encoder,utilization.decoder --format=csv,noheader,nounits 2>$null
        if ($combinedOut) {
            $parts = $combinedOut -split ',\s*'
            if ($parts.Count -ge 5) {
                try {
                    $gpuPower  = [double]$parts[0].Trim()
                    $gpuSmUtil  = [double]$parts[1].Trim()
                    $gpuMemUtil = [double]$parts[2].Trim()
                    $gpuEncUtil = [double]$parts[3].Trim()
                    $gpuDecUtil = [double]$parts[4].Trim()
                } catch {}
            }
        }

        # Fallbacks if batched query failed
        if ($gpuPower -le 0) { $gpuPower = $this.GetGpuPower() }

        # Track total GPU energy (Joules)
        $this.GpuEnergyJoules += ($gpuPower * $dt)

        # Calculate total GPU weighted utilization
        # GPU_weighted = a * GPU_SM + b * GPU_Mem + c * GPU_Enc + d * GPU_Dec
        $gpuWeightedTotal = $this.WeightSm * $gpuSmUtil + `
            $this.WeightMemory * $gpuMemUtil + `
            $this.WeightEncoder * $gpuEncUtil + `
            $this.WeightDecoder * $gpuDecUtil

        $gpuActivePower = $gpuPower - $this.GpuIdlePower
        if ($gpuActivePower -lt 0) { $gpuActivePower = 0 }

        # Get running processes with their utilizations
        $processes = $this.GetGpuProcesses()
        $currentProcessCount = $processes.Count

        # No activity, split idle power equally
        if ($gpuWeightedTotal -le 0) {
            $idleCount = [int]$this.IdleProcessCount
            $idlePowerPerProc = if ($idleCount -gt 0) { $this.GpuIdlePower / $idleCount } else { 0 }
            $numActive = $currentProcessCount - $idleCount
            $activePowerPerProc = if ($numActive -gt 0) { $gpuActivePower / $numActive } else { 0 }
            foreach ($p in $processes) {
                $pidKey = [string]$p.ProcessId
                if ($this.GpuIdleProcesses.ContainsKey($pidKey)) {
                    $energyJ = $idlePowerPerProc * $dt
                }
                else {
                    $energyJ = $activePowerPerProc * $dt
                }
                if (-not $this.ProcessEnergyJoules.ContainsKey($pidKey)) { $this.ProcessEnergyJoules[$pidKey] = 0.0 }
                $this.ProcessEnergyJoules[$pidKey] += $energyJ
            }
            return
        }

        # Calculate per-process total GPU weighted utilization
        $processWeightTotal = 0.0
        foreach ($p in $processes) {
            $weighted = ($this.WeightSm * $p.SmUtil) + 
                       ($this.WeightMemory * $p.MemUtil) + 
                       ($this.WeightEncoder * $p.EncUtil) + 
                       ($this.WeightDecoder * $p.DecUtil)
            $p | Add-Member -NotePropertyName 'WeightedUtil' -NotePropertyValue $weighted -Force
            $processWeightTotal += $weighted
        }
        
        $gpuTotalProcessUtil = $processWeightTotal / $gpuWeightedTotal
        $gpuTotalActiveProcessPower = $gpuTotalProcessUtil * $gpuActivePower
        $gpuExcessPower = $gpuActivePower - $gpuTotalActiveProcessPower

        # Compact / named-variable implementation of the attribution formula.
        # Definitions used below (matching derived math):
        #   P_gpu      = $gpuPower
        #   P_idle     = $this.GpuIdlePower
        #   P_act      = max(0, P_gpu - P_idle) => $gpuActivePower
        #   a,b,c,d    = weights ($this.WeightA..D)
        #   W_i        = process weighted util = a*u_SM + b*u_Mem + c*u_Enc + d*u_Dec
        #   W_tot      = sum_i W_i => $processWeightTotal
        #   G_wt       = a*GPU_SM + b*GPU_Mem + c*GPU_Enc + d*GPU_Dec => $gpuWeightedTotal
        #   P_proc_act_total = ($processWeightTotal / $gpuWeightedTotal) * P_act => $gpuTotalActiveProcessPower
        #   P_excess   = P_act - P_proc_act_total => $gpuExcessPower
        #   P_idle_pwr = P_idle / N_idle  (if N_idle>0)
        #   P_resid_per_proc = P_excess / N_cur  (if N_cur>0)
        # Final per-process power P_i (piecewise):
        #   If process was idle:
        #     if W_i == 0: P_i = P_idle_pwr + P_resid_per_proc
        #     else:         P_i = P_idle_pwr + (W_i/W_tot)*P_proc_act_total + P_resid_per_proc
        #   If process was NOT idle:
        #     if W_i == 0: P_i = P_resid_per_proc
        #     else:         P_i = (W_i/W_tot)*P_proc_act_total + P_resid_per_proc

        $P_idle_pwr = if ($this.IdleProcessCount -gt 0) { $this.GpuIdlePower / $this.IdleProcessCount } else { 0 }
        # Use canonical GPU variables directly; compute residual per-proc from gpuExcessPower
        $P_residual_per_proc = if ($currentProcessCount -gt 0) { $gpuExcessPower / $currentProcessCount } else { 0 }

        $sampleAttributedPower = 0.0
        $samplePerProcess = @{}
        foreach ($p in $processes) {
            $pidKey = [string]$p.ProcessId
            $W_i = $p.WeightedUtil

            # Safe guard for zero total weight
            if ($processWeightTotal -le 0) { $fraction = 0 } else { $fraction = $W_i / $processWeightTotal }

            if ($this.GpuIdleProcesses.ContainsKey($pidKey)) {
                if ($W_i -le 0) {
                    $power = $P_idle_pwr + $P_residual_per_proc
                }
                else {
                    $power = $P_idle_pwr + ($fraction * $gpuTotalActiveProcessPower) + $P_residual_per_proc
                }
            }
            else {
                if ($W_i -le 0) {
                    $power = $P_residual_per_proc
                }
                else {
                    $power = ($fraction * $gpuTotalActiveProcessPower) + $P_residual_per_proc
                }
            }

            $energyJ = $power * $dt
            if (-not $this.ProcessEnergyJoules.ContainsKey($pidKey)) { $this.ProcessEnergyJoules[$pidKey] = 0.0 }
            $this.ProcessEnergyJoules[$pidKey] += $energyJ
            $sampleAttributedPower += $power
            $samplePerProcess[$pidKey] = [PSCustomObject]@{
                PID = $pidKey
                ProcessName = $p.ProcessName
                PowerW = [math]::Round($power,4)
                EnergyJ = [math]::Round($energyJ,6)
                WeightedUtil = $W_i
                IsIdle = [bool]$this.GpuIdleProcesses.ContainsKey($pidKey)
            }
        }

        # Record sample
        $timestamp = ($this.GetGpuTime()).ToString('yyyy/MM/dd HH:mm:ss.fff')
        $sampleResidualPower = $gpuPower - $sampleAttributedPower
        [void]$this.Samples.Add([PSCustomObject]@{
                Timestamp    = $timestamp
                PowerW       = $gpuPower
                GpuSmUtil    = $gpuSmUtil
                GpuMemUtil   = $gpuMemUtil
                GpuEncUtil   = $gpuEncUtil
                GpuDecUtil   = $gpuDecUtil
                ProcessCount = $currentProcessCount
            AttributedPowerW = [math]::Round($sampleAttributedPower,4)
            ResidualPowerW   = [math]::Round($sampleResidualPower,4)
            PerProcess = $samplePerProcess
            })
    }

    [void] Run([int]$Duration) {
        # Wait for user to start workload
        Write-Host ""; Write-Log "READY: Start the workload now (e.g., compute or memory workload)."
        Write-Host "Press ENTER when the workload is running and you want to begin sampling..."
        [void][System.Console]::ReadLine()

        $now = [datetime]($this.GetGpuTime())
        $end = $now.AddSeconds([double]$Duration)
        Write-Log "Monitoring GPU for $Duration seconds..."
        Write-Log "End time (GPU): $end"

        # Check what processes we can see
        Write-Host ""; Write-Log "Checking for GPU processes..."
        $runningProcesses = $this.GetGpuProcesses()
        if ($runningProcesses.Count -eq 0) {
            Write-Host "ERROR: No GPU processes found at all!" -ForegroundColor Red
            return
        }
        
        Write-Log "Found $($runningProcesses.Count) GPU process(es):"
        foreach ($p in $runningProcesses) {
            Write-Host ("  PID: {0,-7} {1,-25} SM%:{2,-5} Mem%:{3,-5} Enc%:{4,-5} Dec%:{5,-5}" -f `
                $p.ProcessId, $p.ProcessName, $p.SmUtil, $p.MemUtil, $p.EncUtil, $p.DecUtil)
        }
        Write-Host ""

        $sampleCount = 0
        while (($this.GetGpuTime()) -lt $end) {
            $this.Sample()
            $sampleCount++
            Start-Sleep -Milliseconds $this.SampleIntervalMs
        }

        Write-Log "Collected $sampleCount samples"
        $this.Report($Duration, $this.TargetProcess)
    }

    [void] Report([double]$Duration, [string]$TargetProcess) {
        Write-Host "`n==== GPU POWER SUMMARY ===="
        
        # Total GPU stats
        $avgGpuPower = $this.GpuEnergyJoules / $Duration
        Write-Host ("Total GPU - Accumulative: {0,8:F2} J  |  Average: {1,6:F2} W" -f $this.GpuEnergyJoules, $avgGpuPower)
        
        Write-Host "`n==== PROCESS POWER ATTRIBUTION (Multi-Metric) ===="

        if ($this.ProcessEnergyJoules.Count -eq 0) {
            Write-Host "No process energy data recorded."
            return
        }

        # Materialize entries so we can iterate multiple times
        $entries = @()
        $this.ProcessEnergyJoules.GetEnumerator() | ForEach-Object { $entries += $_ }

        # Cache PID -> ProcessName via a single Get-Process call to avoid repeated queries
        $allPids = $entries | ForEach-Object { [int]$_.Key } | Sort-Object -Unique
        $pidToName = @{}
        if ($allPids.Count -gt 0) {
            try {
                $procObjs = Get-Process -Id $allPids -ErrorAction SilentlyContinue
                foreach ($pr in $procObjs) { $pidToName[[int]$pr.Id] = $pr.ProcessName }
            }
            catch { }
        }

        if ($TargetProcess) {
            # If target is numeric PID, filter by PID keys; otherwise filter by process name
            if ($TargetProcess -match '^\d+$') {
                $entries = $entries | Where-Object { $_.Key -eq ([string]$TargetProcess) }
            }
            else {
                $entries = $entries | Where-Object {
                    $pidNum = [int]$_.Key
                    $name = $null
                    if ($pidToName.ContainsKey($pidNum)) { $name = $pidToName[$pidNum] }
                    return $name -and ($name -like "*${TargetProcess}*")
                }
            }
            if (-not $entries) { Write-Host "No matching GPU processes recorded."; return }
        }

        $sumTotalProcessesEnergy = ($entries | Measure-Object Value -Sum).Sum

        # Sanity check: summed per-process energy should approximately equal total GPU energy
        if ($sumTotalProcessesEnergy -gt 0) {
            $diff = [math]::Abs([double]$this.GpuEnergyJoules - $sumTotalProcessesEnergy)
            $diffPct = if ($this.GpuEnergyJoules -gt 0) { 100.0 * $diff / $this.GpuEnergyJoules } else { 0 }
            if ($diff -gt 0.5) {
                if ($this.AutoScaleEnabled) {
                    Write-Host ("WARNING: Sum of per-process energy {0:F2} J differs from total GPU energy {1:F2} J. Scaling per-process values to match total." -f $sumTotalProcessesEnergy, [double]$this.GpuEnergyJoules)
                    $scale = if ($sumTotalProcessesEnergy -ne 0) { [double]$this.GpuEnergyJoules / $sumTotalProcessesEnergy } else { 1.0 }
                    # Update both the internal hashtable and the local entries array
                    foreach ($e in $entries) {
                        $k = $e.Key
                        $v = [double]$e.Value
                        $newV = $v * $scale
                        $this.ProcessEnergyJoules[$k] = $newV
                        $e.Value = $newV
                    }
                    # recompute total after scaling
                    $sumTotalProcessesEnergy = ($entries | Measure-Object Value -Sum).Sum
                }
                else {
                    Write-Host ("DIAGNOSTIC: Measured total {0:F2} J vs summed attributed {1:F2} J -> Diff: {2:F2} J ({3:F1}%)" -f $this.GpuEnergyJoules, $sumTotalProcessesEnergy, $diff, $diffPct)
                        if ($this.DiagnosticsEnabled -and $this.DiagnosticsOutputPath) {
                        try {
                            $csvPath = $this.DiagnosticsOutputPath
                            $lines = @()
                            $lines += "Timestamp,PowerW,AttributedPowerW,ResidualPowerW,ProcessCount"
                            foreach ($s in $this.Samples) {
                                $lines += ("{0},{1},{2},{3},{4}" -f $s.Timestamp, $s.PowerW, $s.AttributedPowerW, $s.ResidualPowerW, $s.ProcessCount)
                            }
                            $lines | Out-File -FilePath $csvPath -Encoding UTF8
                            Write-Host "Diagnostics written to $csvPath"

                            # Also write per-sample per-process detailed CSV
                            $csvPathProcesses = $csvPath + ".processes.csv"
                            $lines2 = @()
                            $lines2 += "Timestamp,PID,ProcessName,PowerW,EnergyJ,WeightedUtil,IsIdle"
                            foreach ($s in $this.Samples) {
                                if ($s.PerProcess) {
                                    foreach ($kv in $s.PerProcess.GetEnumerator()) {
                                        $pp = $kv.Value
                                        $lines2 += ("{0},{1},{2},{3},{4},{5},{6}" -f $s.Timestamp, $pp.PID, ($pp.ProcessName -replace ',',' '), $pp.PowerW, $pp.EnergyJ, $pp.WeightedUtil, $pp.IsIdle)
                                    }
                                }
                            }
                            $lines2 | Out-File -FilePath $csvPathProcesses -Encoding UTF8
                            Write-Host "Per-process diagnostics written to $csvPathProcesses"
                        }
                        catch {
                            Write-Host "Failed to write diagnostics to $($this.DiagnosticsOutputPath): $_" -ForegroundColor Yellow
                        }
                    }
                }
            }
        }

        # Aggregate by process name to handle multiple PIDs with same name
        $agg = @{}
        foreach ($e in $entries) {
            $pidNum = [int]$e.Key
            $pname = $null
            if ($pidToName.ContainsKey($pidNum)) { $pname = $pidToName[$pidNum] }
            if (-not $pname) { $pname = "PID $pidNum" }
            if (-not $agg.ContainsKey($pname)) { $agg[$pname] = @{ Energy = 0.0; Pids = @() } }
            $agg[$pname].Energy += $e.Value
            $agg[$pname].Pids += $pidNum
        }

        foreach ($k in $agg.GetEnumerator() | Sort-Object @{Expression={$_.Value.Energy};Descending=$true}) {
            $name = $k.Key
            $energy = $k.Value.Energy
            $avgPower = $energy / $Duration
            # When a TargetProcess is specified show percent of total GPU energy,
            # otherwise show percent relative to the sum of recorded process energies.
            if ($this.GpuEnergyJoules -gt 0) {
                if ($TargetProcess) { $pct = 100 * $energy / $this.GpuEnergyJoules } else { $pct = if ($sumTotalProcessesEnergy -gt 0) { 100 * $energy / $sumTotalProcessesEnergy } else { 0 } }
            }
            else { $pct = 0 }
            $pidList = ($k.Value.Pids -join ',')
            Write-Host ("{0,-30} Accumulative: {1,8:F2} J  |  Average: {2,6:F2} W  ({3,5:F1}%)  PIDs: {4}" -f $name, $energy, $avgPower, $pct, $pidList)
        }

        # Show percentage of total GPU power for target if requested
        if ($TargetProcess -and $sumTotalProcessesEnergy -gt 0) {
            $targetPct = if ($this.GpuEnergyJoules -gt 0) { ($sumTotalProcessesEnergy / $this.GpuEnergyJoules) * 100 } else { 0 }
            Write-Host "`nTarget process used $($targetPct.ToString('F1'))% of total GPU power"
        }
    }
}

# Main execution
    try {
    $monitor = [GPUProcessMonitor]::new($Process, $SampleInterval, $WeightSM, $WeightMem, $WeightEnc, $WeightDec, $AutoScale.IsPresent, $Diagnostics.IsPresent, $DiagnosticsOutput)
    $monitor.Run($Duration)
}
catch {
    Write-Error "Error: $_"
    exit 1
}