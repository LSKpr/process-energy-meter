<#
.SYNOPSIS
    GPU Power Monitor using nvidia-smi
.DESCRIPTION
    Monitors GPU power consumption and attributes it to processes
.PARAMETER Process
    Target process name to track (substring match)
.PARAMETER Duration
    Duration in seconds (default: 60)
.PARAMETER SampleInterval
    Sample interval in milliseconds (default: 10)
#>

param(
    [string]$Process = $null,
    [int]$Duration = 60,
    [int]$SampleInterval = 10
)

class GPUPowerMonitor {
    [string]$TargetProcess
    [int]$SampleIntervalMs
    [double]$SampleIntervalSec
    [System.Collections.ArrayList]$Samples
    [hashtable]$ProcessEnergy
    [string]$GpuName

    GPUPowerMonitor([string]$targetProcess, [int]$sampleIntervalMs) {
        $this.TargetProcess = $targetProcess
        $this.SampleIntervalMs = $sampleIntervalMs
        $this.SampleIntervalSec = $sampleIntervalMs / 1000.0
        $this.Samples = [System.Collections.ArrayList]::new()
        $this.ProcessEnergy = @{}

        # Get GPU name
        $gpuInfo = nvidia-smi --query-gpu=name --format=csv,noheader 2>$null
        if ($LASTEXITCODE -ne 0) {
            Write-Error "nvidia-smi not found or failed to execute"
            exit 1
        }
        $this.GpuName = $gpuInfo.Trim()
        Write-Host "Monitoring GPU: $($this.GpuName)"

        # If target process is given, check if it's running
        if ($this.TargetProcess) {
            if (-not $this.IsProcessRunning()) {
                Write-Host "ERROR: Target process '$($this.TargetProcess)' not found running on GPU.`n" -ForegroundColor Red
                $this.ListGpuProcesses()
                exit 1
            } else {
                Write-Host "Tracking target process: $($this.TargetProcess)"
            }
        }
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

    [array] GetGpuProcesses() {
        $processes = @()
        
        # Method 1: Parse nvidia-smi pmon output (shows all GPU processes)
        try {
            $pmonOutput = nvidia-smi pmon -c 1 2>$null
            if ($pmonOutput) {
                $lines = $pmonOutput -split "`n"
                foreach ($line in $lines) {
                    # Skip header lines and empty lines
                    if ($line -match '^\s*#' -or $line -match '^-+' -or [string]::IsNullOrWhiteSpace($line)) {
                        continue
                    }
                    
                    # Parse: gpu pid type sm mem enc dec command
                    if ($line -match '^\s*(\d+)\s+(\d+)\s+(\w+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(.+?)\s*$') {
                        $processId = [int]$Matches[2]
                        $memUtil = $Matches[5]
                        $command = $Matches[8].Trim()
                        
                        # Try to get actual memory usage from nvidia-smi
                        $memMB = $this.GetProcessMemory($processId)
                        
                        $processes += [PSCustomObject]@{
                            ProcessId = $processId
                            ProcessName = $command
                            UsedMemoryMB = $memMB
                        }
                    }
                }
            }
        } catch {
            # Fallback method
        }
        
        # Method 2: Fallback - use Windows processes and match with GPU memory
        if ($processes.Count -eq 0) {
            # Get total GPU memory usage per process using WMI/Get-Process
            $gpuProcessIds = $this.GetGpuProcessIdsFromSmi()
            foreach ($processId in $gpuProcessIds) {
                try {
                    $proc = Get-Process -Id $processId -ErrorAction SilentlyContinue
                    if ($proc) {
                        $memMB = $this.GetProcessMemory($processId)
                        $processes += [PSCustomObject]@{
                            ProcessId = $processId
                            ProcessName = $proc.ProcessName
                            UsedMemoryMB = $memMB
                        }
                    }
                } catch {
                    continue
                }
            }
        }
        
        return $processes
    }

    [double] GetProcessMemory([int]$processId) {
        # Try to query process-specific GPU memory
        try {
            $output = nvidia-smi --query-compute-apps=pid,used_memory --format=csv,noheader,nounits 2>$null
            if ($output) {
                foreach ($line in $output -split "`n") {
                    if ([string]::IsNullOrWhiteSpace($line)) { continue }
                    $parts = $line -split ',\s*'
                    if ($parts.Count -ge 2 -and [int]$parts[0].Trim() -eq $processId) {
                        $mem = $parts[1].Trim()
                        if ($mem -ne '[N/A]' -and $mem -match '^\d+') {
                            return [double]$mem
                        }
                    }
                }
            }
        } catch {}
        
        # Default fallback - estimate based on total GPU memory
        return 100.0  # Default 100MB if we can't determine
    }

    [array] GetGpuProcessIdsFromSmi() {
        $pids = @()
        try {
            # Use pmon to get all GPU process IDs
            $pmonOutput = nvidia-smi pmon -c 1 2>$null
            if ($pmonOutput) {
                foreach ($line in $pmonOutput -split "`n") {
                    if ($line -match '^\s*\d+\s+(\d+)') {
                        $pids += [int]$Matches[1]
                    }
                }
            }
        } catch {}
        return $pids | Select-Object -Unique
    }

    [void] ListGpuProcesses() {
        $processes = $this.GetGpuProcesses()
        Write-Host "GPU processes currently running:"
        if ($processes.Count -eq 0) {
            Write-Host "  None"
        } else {
            foreach ($proc in $processes) {
                Write-Host "  PID $($proc.ProcessId): $($proc.ProcessName) ($($proc.UsedMemoryMB) MB)"
            }
        }
    }

    [void] Sample() {
        # Get power draw in watts
        $powerOutput = nvidia-smi --query-gpu=power.draw --format=csv,noheader,nounits 2>$null
        $powerW = 0.0
        if ($powerOutput -and $powerOutput.Trim() -ne '[N/A]') {
            try {
                $powerW = [double]$powerOutput.Trim()
            } catch {
                $powerW = 0.0
            }
        }

        # Get GPU utilization
        $utilOutput = nvidia-smi --query-gpu=utilization.gpu --format=csv,noheader,nounits 2>$null
        $util = 0.0
        if ($utilOutput -and $utilOutput.Trim() -ne '[N/A]') {
            try {
                $util = [double]$utilOutput.Trim()
            } catch {
                $util = 0.0
            }
        }

        # Get running processes
        $processes = $this.GetGpuProcesses()

        # Calculate total memory
        $totalMem = ($processes | Measure-Object -Property UsedMemoryMB -Sum).Sum
        if ($totalMem -eq 0) { $totalMem = 1 }

        # Record sample
        $timestamp = (Get-Date).ToUniversalTime().ToString("o")
        [void]$this.Samples.Add([PSCustomObject]@{
            Timestamp = $timestamp
            PowerW = $powerW
            GpuUtil = $util
        })

        # Attribute power to processes
        foreach ($proc in $processes) {
            $processName = $proc.ProcessName

            # Track only target process if specified
            if ($this.TargetProcess -and $processName -notlike "*$($this.TargetProcess)*") {
                continue
            }

            $memFrac = $proc.UsedMemoryMB / $totalMem
            $procPower = $powerW * $memFrac
            $energyJ = $procPower * $this.SampleIntervalSec

            if (-not $this.ProcessEnergy.ContainsKey($processName)) {
                $this.ProcessEnergy[$processName] = 0.0
            }
            $this.ProcessEnergy[$processName] += $energyJ
        }
    }

    [void] Run([int]$durationSec) {
        $startTime = Get-Date
        $endTime = $startTime.AddSeconds($durationSec)

        Write-Host "`nMonitoring for $durationSec seconds (Ctrl+C to stop early)...`n"

        try {
            while ((Get-Date) -lt $endTime) {
                $this.Sample()
                Start-Sleep -Milliseconds $this.SampleIntervalMs
            }
        } catch {
            Write-Host "`nMonitoring interrupted" -ForegroundColor Yellow
        }

        $actualDuration = ((Get-Date) - $startTime).TotalSeconds
        $this.Report($actualDuration)
    }

    [void] Report([double]$duration) {
        # Calculate total energy
        $totalEnergyJ = ($this.Samples | Measure-Object -Property PowerW -Sum).Sum * $this.SampleIntervalSec
        $totalEnergyKwh = $totalEnergyJ / 3600000

        Write-Host "`n========== GPU RESULTS =========="
        Write-Host ("Duration: {0:F2} s" -f $duration)
        Write-Host ("Total GPU Energy: {0:F6} kWh" -f $totalEnergyKwh)

        Write-Host "`n====== PER-PROCESS ENERGY ======"
        if ($this.ProcessEnergy.Count -eq 0) {
            Write-Host "  No process data collected"
        } else {
            foreach ($entry in $this.ProcessEnergy.GetEnumerator() | Sort-Object Value -Descending) {
                $name = $entry.Key
                $energyJ = $entry.Value
                $energyKwh = $energyJ / 3600000
                $pct = if ($totalEnergyJ -gt 0) { ($energyJ / $totalEnergyJ * 100) } else { 0 }
                Write-Host ("{0,-30} {1:F6} kWh ({2:F1}%)" -f $name, $energyKwh, $pct)
            }
        }

        $this.SaveCsv()
    }

    [void] SaveCsv() {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $filename = "gpu_log_$timestamp.csv"

        $this.Samples | Export-Csv -Path $filename -NoTypeInformation

        Write-Host "`nSaved log to $filename"
    }
}

# Main execution
try {
    $monitor = [GPUPowerMonitor]::new($Process, $SampleInterval)
    $monitor.Run($Duration)
} catch {
    Write-Error "Error: $_"
    exit 1
}