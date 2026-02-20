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

# run.ps1 — launcher for the producer + UI split
# Usage: Open elevated PowerShell and run: .\run.ps1

# ---------------------------
#region Main Execution
# ---------------------------

# Ensure we run from script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $scriptDir

# Dot-source logic (producer) first, then the interface (UI)
. .\PowerSampleLogic.ps1
. .\PowerInterface.ps1

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