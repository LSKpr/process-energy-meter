# PCM Driver Installation Script
# Run this as Administrator

$pcmDir = Join-Path $PSScriptRoot "pcm"
$msrSys = Join-Path $pcmDir "MSR.sys"

Write-Host "`nPCM Driver Installation" -ForegroundColor Cyan
Write-Host "=" * 50 -ForegroundColor Cyan

# Check admin
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "`nERROR: Must run as Administrator!" -ForegroundColor Red
    exit 1
}

# Check if MSR.sys exists
if (-not (Test-Path $msrSys)) {
    Write-Host "`nERROR: MSR.sys not found at $msrSys" -ForegroundColor Red
    exit 1
}

Write-Host "`n[*] Installing PCM MSR driver..." -ForegroundColor Yellow

# Remove existing service if present
$existingService = Get-Service -Name "PCM-MSR" -ErrorAction SilentlyContinue
if ($existingService) {
    Write-Host "[*] Removing existing PCM-MSR service..." -ForegroundColor Gray
    & sc.exe delete "PCM-MSR" | Out-Null
    Start-Sleep -Seconds 1
}

# Create the service
Write-Host "[*] Creating PCM-MSR service..." -ForegroundColor Gray
$result = & sc.exe create "PCM-MSR" binPath= $msrSys type= kernel start= demand error= normal DisplayName= "Intel PCM MSR Driver"

if ($LASTEXITCODE -ne 0) {
    Write-Host "`n[X] Failed to create service. Output:" -ForegroundColor Red
    Write-Host $result -ForegroundColor Gray
    exit 1
}

Write-Host "[OK] Service created" -ForegroundColor Green

# Start the service
Write-Host "[*] Starting PCM-MSR service..." -ForegroundColor Gray
$result = & sc.exe start "PCM-MSR"

if ($LASTEXITCODE -ne 0) {
    Write-Host "`n[!] Service start returned code: $LASTEXITCODE" -ForegroundColor Yellow
    Write-Host $result -ForegroundColor Gray
    
    # Check if it's a driver signature issue
    if ($result -like "*signature*" -or $result -like "*1275*") {
        Write-Host "`n[!] Driver signature issue detected" -ForegroundColor Yellow
        Write-Host "    Windows may be blocking unsigned drivers" -ForegroundColor Gray
        Write-Host "`n    Options:" -ForegroundColor Yellow
        Write-Host "    1. Use LibreHardwareMonitor (RAPL) instead - already works!" -ForegroundColor Green
        Write-Host "    2. Disable driver signature enforcement (advanced, not recommended)" -ForegroundColor Gray
    }
}
else {
    Write-Host "[OK] PCM-MSR service started successfully!" -ForegroundColor Green
}

# Verify service status
Start-Sleep -Seconds 1
$service = Get-Service -Name "PCM-MSR" -ErrorAction SilentlyContinue
if ($service) {
    Write-Host "`n[*] Service Status: $($service.Status)" -ForegroundColor Cyan
    
    if ($service.Status -eq "Running") {
        Write-Host "`n[OK] PCM driver is ready!" -ForegroundColor Green
        Write-Host "`nYou can now run:" -ForegroundColor Yellow
        Write-Host "  cd `"$pcmDir`"" -ForegroundColor Gray
        Write-Host "  .\pcm-sensor-server.exe -p 9738" -ForegroundColor Gray
    }
    else {
        Write-Host "`n[!] Service not running - there may be driver compatibility issues" -ForegroundColor Yellow
        Write-Host "    Recommendation: Use LibreHardwareMonitor (RAPL) instead" -ForegroundColor Green
    }
}

Write-Host ""
