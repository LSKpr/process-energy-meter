# Download and use LibreHardwareMonitor for RAPL access

$libPath = Join-Path $PSScriptRoot "LibreHardwareMonitorLib.dll"

if (-not (Test-Path $libPath)) {
    Write-Host "LibreHardwareMonitor library not found." -ForegroundColor Yellow
    Write-Host "Downloading from GitHub..." -ForegroundColor Cyan
    
    # Get the latest release
    $releasesUrl = "https://api.github.com/repos/LibreHardwareMonitor/LibreHardwareMonitor/releases/latest"
    $release = Invoke-RestMethod -Uri $releasesUrl
    $asset = $release.assets | Where-Object { $_.name -like "*.zip" } | Select-Object -First 1
    
    if (-not $asset) {
        Write-Host "Could not find download asset" -ForegroundColor Red
        Write-Host "`nManual steps:" -ForegroundColor Yellow
        Write-Host "1. Download from: https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases" -ForegroundColor White
        Write-Host "2. Extract and copy LibreHardwareMonitorLib.dll to: $PSScriptRoot" -ForegroundColor White
        exit 1
    }
    
    $downloadUrl = $asset.browser_download_url
    $zipPath = Join-Path $PSScriptRoot "LibreHardwareMonitor.zip"
    
    try {
        Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -UseBasicParsing
        Expand-Archive -Path $zipPath -DestinationPath $PSScriptRoot -Force
        
        # Find the DLL
        $dllFound = Get-ChildItem -Path $PSScriptRoot -Recurse -Filter "LibreHardwareMonitorLib.dll" | Select-Object -First 1
        
        if ($dllFound) {
            Copy-Item $dllFound.FullName -Destination $libPath
            Write-Host "Library downloaded successfully!" -ForegroundColor Green
        }
        
        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "Failed to download: $_" -ForegroundColor Red
        Write-Host "`nManual steps:" -ForegroundColor Yellow
        Write-Host "1. Download from: https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases" -ForegroundColor White
        Write-Host "2. Extract and copy LibreHardwareMonitorLib.dll to: $PSScriptRoot" -ForegroundColor White
        exit 1
    }
}

# Load the library
try {
    Add-Type -Path $libPath
    Write-Host "LibreHardwareMonitor loaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to load library: $_" -ForegroundColor Red
    exit 1
}

# Test RAPL reading
Write-Host "`nTesting CPU power reading..." -ForegroundColor Cyan

$computer = New-Object LibreHardwareMonitor.Hardware.Computer
$computer.IsCpuEnabled = $true
$computer.Open()

Write-Host "Computer initialized" -ForegroundColor Green

foreach ($hardware in $computer.Hardware) {
    if ($hardware.HardwareType -eq [LibreHardwareMonitor.Hardware.HardwareType]::Cpu) {
        Write-Host "`nCPU Found: $($hardware.Name)" -ForegroundColor Yellow
        
        $hardware.Update()
        
        Write-Host "`nAll sensors:" -ForegroundColor Cyan
        foreach ($sensor in $hardware.Sensors) {
            Write-Host "  [$($sensor.SensorType)] $($sensor.Name): $($sensor.Value)" -ForegroundColor White
        }
        
        Write-Host "`nPower sensors:" -ForegroundColor Green
        $powerSensors = $hardware.Sensors | Where-Object { $_.SensorType -eq [LibreHardwareMonitor.Hardware.SensorType]::Power }
        if ($powerSensors) {
            foreach ($sensor in $powerSensors) {
                Write-Host "  $($sensor.Name): $($sensor.Value) W" -ForegroundColor Green
            }
        } else {
            Write-Host "  No power sensors found" -ForegroundColor Red
        }
    }
}

$computer.Close()

Write-Host "`nIf you see CPU power readings above, LibreHardwareMonitor is working!" -ForegroundColor Cyan
Write-Host "You can integrate this into the power meter script." -ForegroundColor Cyan
