<#
Usage example:
For generall monitoring of all GPU processes:
.\gpu_complete_clean.ps1
For specific process by name or PID (e.g., "python" or "1234""):
.\gpu_complete_clean.ps1 -Process "python"
Other flags of intereset:
    -Duration: how long to monitor in seconds (default 60)
    -SampleInterval: how often to sample in ms (default 100)
    -WeightSM, WeightMem, WeightEnc, WeightDec: weights for the attribution formula (defaults: 1.0, 0.5, 0.25, 0.15)
    -AutoScale: if set, will scale per-process values to match total GPU energy if there is a significant difference (e.g., due to measurement error)
    -DiagnosticsOutput: base path for writing diagnostics CSV files (default "gpu_diagnostics")
#>

param(
    [string]$Process = $null,
    [int]$Duration = 60,
    [int]$SampleInterval = 100,
    [double]$WeightSM = 1.0,
    [double]$WeightMem = 0.5,
    [double]$WeightEnc = 0.25,
    [double]$WeightDec = 0.15,
    [switch]$AutoScale,
    [string]$DiagnosticsOutput = "gpu_diagnostics"
)

function Write-Log($msg) {
    $t = Get-Date
    $tStr = $t.ToString('yyyy/MM/dd HH:mm:ss.fff')
    Write-Host "[$tStr] $msg"
}

class GPUProcessMonitor {
    [string]$TargetProcess
    [int]$SampleIntervalMs
    [double]$WeightSm      # SM weight
    [double]$WeightMemory  # Memory weight
    [double]$WeightEncoder # Encoder weight
    [double]$WeightDecoder # Decoder weight

    [System.Collections.Generic.List[object]]$Samples
    [datetime]$LastSampleTime

    [hashtable]$ProcessEnergyJoules
    [string]$GpuName
    [double]$GpuEnergyJoules
    [double]$GpuIdlePower
    [hashtable]$GpuIdleProcesses
    [int]$IdleProcessCount
    
    [bool]$AutoScaleEnabled
    [string]$DiagnosticsOutputPath

    GPUProcessMonitor([string]$targetProcess, [int]$SampleIntervalMs, [double]$wSM, [double]$wMem, [double]$wEnc, [double]$wDec, [bool]$autoScale, [string]$diagPath) {
        $this.TargetProcess = $targetProcess
        $this.SampleIntervalMs = $SampleIntervalMs
        $this.WeightSm = $wSM
        $this.WeightMemory = $wMem
        $this.WeightEncoder = $wEnc
        $this.WeightDecoder = $wDec
        
        $this.Samples = [System.Collections.Generic.List[object]]::new()
        $this.LastSampleTime = Get-Date

        $this.ProcessEnergyJoules = @{}
        $this.GpuEnergyJoules = 0.0
        $this.GpuIdleProcesses = @{}
        
        $this.AutoScaleEnabled = $autoScale
        $this.DiagnosticsOutputPath = $diagPath

        # Get GPU name
        $gpuNameOutput = nvidia-smi --query-gpu=name --format=csv,noheader 2>$null
        $this.GpuName = $gpuNameOutput.Trim()
        Write-Log "Monitoring GPU: $($this.GpuName)"
        Write-Host "Power attribution weights: SM=$wSM, Mem=$wMem, Enc=$wEnc, Dec=$wDec"

        # Measure idle metrics
        $this.MeasureIdleMetrics()
    }

    [void] MeasureIdleMetrics() {
        Write-Log "Measuring idle power (please ensure GPU is idle; close heavy apps)."
        # Take multiple samples to get a stable idle measurement using a single batched nvidia-smi call
        $idlePowerSamples = @()
        $idlePowerMin = 0
        $idlePowerMax = 0
        for ($i = 0; $i -lt 10; $i++) {
            $powerOutput = nvidia-smi --query-gpu=power.draw.instant --format=csv,noheader,nounits 2>$null
            if ($powerOutput -and $powerOutput.Trim() -ne '[N/A]') {
                try {
                    $idlePowerSamples += [double]$powerOutput.Trim()
                }
                catch {}
            }
            Start-Sleep -Milliseconds $this.SampleIntervalMs
        }

        if ($idlePowerSamples.Count -gt 0) {
            $measuredIdlePower = ($idlePowerSamples | Measure-Object -Average).Average
            $idlePowerMin = $idlePowerSamples | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
            $idlePowerMax = $idlePowerSamples | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
            $this.GpuIdlePower = $measuredIdlePower
        }

        # Record which processes were running during idle measurement
        $idleProcesses = $this.GetGpuProcesses()
        $this.IdleProcessCount = $idleProcesses.Count
        foreach ($proc in $idleProcesses) {
            $this.GpuIdleProcesses[[string]$proc.ProcessId] = $true
        }

        Write-Log (("Idle GPU power measured ({3} samples): {0:N2}W [Min: {1:N2}W  Max:{2:N2}W] with {4} Processes" -f `
                    $this.GpuIdlePower, $idlePowerMin, $idlePowerMax, $idlePowerSamples.Count, $this.IdleProcessCount))
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

        # Build PID list and query Windows processes once to get names
        $pids = $parsed | Select-Object -ExpandProperty ProcessId -Unique
        $procNameMap = @{}
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

            $processes += [PSCustomObject]@{
                ProcessId   = $entry.ProcessId
                ProcessName = $processName
                SmUtil      = $entry.SmUtil
                MemUtil     = $entry.MemUtil
                EncUtil     = $entry.EncUtil
                DecUtil     = $entry.DecUtil
            }
        }

        return $processes
    }

    # GetProcessMemory is no longer used; memory is obtained in GetGpuProcesses() in a batched fashion.

    [double] ParseUtil([string]$value) {
        if ([string]::IsNullOrWhiteSpace($value) -or $value -eq '-') { return 0.0 }
        if ($value -match '([-+]?[0-9]*\.?[0-9]+)') { 
            try { 
                return [double]$matches[1] 
            }
            catch { } 
        }
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
        $gpuPower = 0.0
        $gpuSmUtil = 0.0 
        $gpuMemUtil = 0.0
        $gpuEncUtil = 0.0
        $gpuDecUtil = 0.0
        $timestamp = "-"
        $combinedOut = nvidia-smi --query-gpu=power.draw.instant,utilization.gpu,utilization.memory,utilization.encoder,utilization.decoder,timestamp --format=csv,noheader,nounits 2>$null
        if ($combinedOut) {
            $parts = $combinedOut -split ',\s*'
            if ($parts.Count -ge 6) {
                try {
                    $gpuPower = [double]$parts[0].Trim()
                    $gpuSmUtil = [double]$parts[1].Trim()
                    $gpuMemUtil = [double]$parts[2].Trim()
                    $gpuEncUtil = [double]$parts[3].Trim()
                    $gpuDecUtil = [double]$parts[4].Trim()
                    $timestamp = $parts[5]
                }
                catch {}
            }
        }

        # Track total GPU energy (Joules)
        $this.GpuEnergyJoules += ($gpuPower * $dt)

        # Calculate total GPU weighted utilization
        # GPU_weighted = a * GPU_SM + b * GPU_Mem + c * GPU_Enc + d * GPU_Dec
        $gpuWeightedTotal = ($this.WeightSm * $gpuSmUtil) +
            ($this.WeightMemory * $gpuMemUtil) +
            ($this.WeightEncoder * $gpuEncUtil) +
            ($this.WeightDecoder * $gpuDecUtil)

        # Get running processes with their utilizations
        $processes = $this.GetGpuProcesses()
        $currentProcessCount = $processes.Count

        # Calculate per-process total GPU weighted utilization
        $processWeightTotal = 0.0
        $inactiveIdleProcessesCount = 0
        foreach ($p in $processes) {
            $weighted = ($this.WeightSm * $p.SmUtil) + 
            ($this.WeightMemory * $p.MemUtil) + 
            ($this.WeightEncoder * $p.EncUtil) + 
            ($this.WeightDecoder * $p.DecUtil)
            $p | Add-Member -NotePropertyName 'WeightedUtil' -NotePropertyValue $weighted -Force
            $processWeightTotal += $weighted

            if ($weighted -le 0 -and $this.GpuIdleProcesses.ContainsKey([string]$p.ProcessId)) {
                $inactiveIdleProcessesCount++;
            }
        }

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

        $gpuActivePower = if (($this.GpuIdlePower) -lt $gpuPower) { $gpuPower - $this.GpuIdlePower } else { 0 }     # Maybe add to recalculate idle power?
        $gpuTotalProcessUtil = if ($gpuWeightedTotal -gt 0) { $processWeightTotal / $gpuWeightedTotal } else { 0 }  # Maybe add to recalculate idle power?
        $gpuTotalActiveProcessPower = $gpuTotalProcessUtil * $gpuActivePower
        $gpuExcessPower = $gpuActivePower - $gpuTotalActiveProcessPower # gpuExcessPower = gpuActivePower if gpuWeightedTotal = 0 or processWeightTotal = 0

        $activeProcessCount = if ($inactiveIdleProcessesCount -eq $currentProcessCount) { $currentProcessCount } else { $currentProcessCount - $inactiveIdleProcessesCount }
        $P_idle_pwr = if ($this.IdleProcessCount -gt 0) { $this.GpuIdlePower / $this.IdleProcessCount } else { 0 }
        $P_residual_per_proc = if ($activeProcessCount -gt 0) { $gpuExcessPower / $activeProcessCount } else { 0 }

        $sampleAttributedPower = 0.0
        $samplePerProcess = @{}
        foreach ($p in $processes) {
            $pidKey = [string]$p.ProcessId
            $W_i = $p.WeightedUtil

            # Safe guard for zero total process weight
            if ($processWeightTotal -le 0) { $fraction = 0 } else { $fraction = $W_i / $processWeightTotal }
            if ($this.GpuIdleProcesses.ContainsKey($pidKey)) {
                if ($W_i -le 0) {
                    $power = if ($inactiveIdleProcessesCount -eq $currentProcessCount) { $P_idle_pwr + $P_residual_per_proc } else { $P_idle_pwr }
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
                PID          = $pidKey
                ProcessName  = $p.ProcessName
                PowerW       = [math]::Round($power, 4)
                EnergyJ      = [math]::Round($energyJ, 6)
                WeightedUtil = $W_i
                IsIdle       = [bool]$this.GpuIdleProcesses.ContainsKey($pidKey)
            }
        }

        # Record sample
        $sampleResidualPower = $gpuPower - $sampleAttributedPower
        [void]$this.Samples.Add([PSCustomObject]@{
                Timestamp        = $timestamp
                PowerW           = $gpuPower
                GpuSmUtil        = $gpuSmUtil
                GpuMemUtil       = $gpuMemUtil
                GpuEncUtil       = $gpuEncUtil
                GpuDecUtil       = $gpuDecUtil
                ProcessCount     = $currentProcessCount
                AttributedPowerW = [math]::Round($sampleAttributedPower, 4)
                ResidualPowerW   = [math]::Round($sampleResidualPower, 4)
                PerProcess       = $samplePerProcess
            })
    }

    [void] Run([int]$Duration) {
        # Wait for user to start workload
        Write-Host ""; Write-Log "READY: Start the workload now (e.g., compute or memory workload)."
        Write-Host "Press ENTER when the workload is running and you want to begin sampling..."
        [void][System.Console]::ReadLine()

        $now = Get-Date
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
        while ((Get-Date) -lt $end) {
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
            if ($this.AutoScaleEnabled) {
                Write-Log ("AUTOSCALE: Measured total {0:F2} J vs summed attributed {1:F2} J -> Diff: {2:F2} J ({3:F1}%)" -f $this.GpuEnergyJoules, $sumTotalProcessesEnergy, $diff, $diffPct)
                if ($diff -gt 0.5) {
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
            }
            Write-Log ("DIAGNOSTIC: Measured total {0:F2} J vs summed attributed {1:F2} J -> Diff: {2:F2} J ({3:F1}%)" -f $this.GpuEnergyJoules, $sumTotalProcessesEnergy, $diff, $diffPct)
            try {
                $csvTime = "_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                $csvPath = $this.DiagnosticsOutputPath + $csvTime
                $lines = @()
                $lines += "Timestamp,PowerW,AttributedPowerW,ResidualPowerW,ProcessCount"
                foreach ($s in $this.Samples) {
                    $lines += ("{0},{1},{2},{3},{4}" -f $s.Timestamp, $s.PowerW, $s.AttributedPowerW, $s.ResidualPowerW, $s.ProcessCount)
                }
                $lines | Out-File -FilePath $csvPath -Encoding UTF8
                Write-Host "Diagnostics written to $csvPath"

                # Also write per-sample per-process detailed CSV
                $csvPathProcesses = $this.DiagnosticsOutputPath + "_processes" + $csvTime
                $lines2 = @()
                $lines2 += "Timestamp,PID,ProcessName,PowerW,EnergyJ,WeightedUtil,IsIdle"
                foreach ($s in $this.Samples) {
                    if ($s.PerProcess) {
                        foreach ($kv in $s.PerProcess.GetEnumerator()) {
                            $pp = $kv.Value
                            $lines2 += ("{0},{1},{2},{3},{4},{5},{6}" -f $s.Timestamp, $pp.PID, ($pp.ProcessName -replace ',', ' '), $pp.PowerW, $pp.EnergyJ, $pp.WeightedUtil, $pp.IsIdle)
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

        foreach ($k in $agg.GetEnumerator() | Sort-Object @{Expression = { $_.Value.Energy }; Descending = $true }) {
            $name = $k.Key
            $energy = $k.Value.Energy
            $avgPower = $energy / $Duration
            # Show percent relative to the sum of recorded process energies.
            if ($this.GpuEnergyJoules -gt 0) {
                $pct = if ($sumTotalProcessesEnergy -gt 0) { 100 * $energy / $sumTotalProcessesEnergy } else { 0 } 
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
    $monitor = [GPUProcessMonitor]::new($Process, $SampleInterval, $WeightSM, $WeightMem, $WeightEnc, $WeightDec, $AutoScale.IsPresent, $DiagnosticsOutput)
    $monitor.Run($Duration)
}
catch {
    Write-Error "Error: $_"
    exit 1
}