# Intel RAPL Setup Guide

## Overview

Intel RAPL (Running Average Power Limit) allows direct reading of CPU power consumption through Model-Specific Registers (MSR). This provides **CPU-only power measurement** instead of total system power.

## Requirements

1. **Intel CPU**: Sandy Bridge (2nd gen) or newer
   - Core i3/i5/i7/i9 from 2011 onwards
   - Check: `wmic cpu get name`

2. **WinRing0 Driver**: Required for MSR access
   - Download from: https://github.com/GermanAizek/WinRing0/releases
   - Or use OpenHardwareMonitor which includes it

3. **Administrator Privileges**: MSR access requires elevation

4. **.NET Framework 4.x**: Usually pre-installed on Windows

## Installation Steps

### Option 1: Using OpenHardwareMonitor (Easiest)

1. Download OpenHardwareMonitor: https://openhardwaremonitor.org/downloads/
2. Extract and run `OpenHardwareMonitor.exe` **as Administrator**
3. Keep it running in the background (minimized is fine)
4. The WinRing0 driver is now loaded and available

### Option 2: Manual WinRing0 Installation

1. Download WinRing0 from GitHub
2. Extract files
3. Copy `WinRing0x64.sys` to `C:\Windows\System32\drivers\`
4. Run as Administrator:
   ```powershell
   sc create WinRing0_1_2_0 binPath= "C:\Windows\System32\drivers\WinRing0x64.sys" type= kernel start= demand
   sc start WinRing0_1_2_0
   ```

## Building the RAPL Library

1. Open PowerShell **as Administrator**
2. Navigate to the project folder:
   ```powershell
   cd "C:\Users\kacpe\Desktop\KTH\AL1523"
   ```
3. Run the build script:
   ```powershell
   .\Build-MsrReader.ps1
   ```
4. Verify `MsrReader.dll` was created

## Running with RAPL Support

```powershell
# Run as Administrator
.\ProcessPowerMeter-Enhanced.ps1
```

The script will:
- ✓ Auto-detect RAPL support
- ✓ Use CPU-only power if available
- ✓ Fall back to system power if RAPL unavailable

## Verification

To check if RAPL is working:

```powershell
Add-Type -Path ".\MsrReader.dll"
$monitor = New-Object PowerMeter.RaplPowerMonitor
$monitor.IsSupported  # Should return True
$monitor.GetPackagePower()  # Should return CPU power in Watts
```

## Troubleshooting

### "RAPL not supported"
- **Check CPU**: Must be Intel Sandy Bridge or newer
- **Install driver**: Follow WinRing0 installation steps above
- **Run as Admin**: Required for MSR access

### "Failed to initialize RAPL"
- Verify WinRing0 driver is loaded: `sc query WinRing0_1_2_0`
- Try running OpenHardwareMonitor first
- Check Windows Event Viewer for driver errors

### "MsrReader.dll not found"
- Run `.\Build-MsrReader.ps1` first
- Verify .NET Framework is installed

### Still not working?
- Use the original script: `.\ProcessPowerMeter.ps1`
- It uses system-wide power measurement (still functional)

## What RAPL Measures

- **CPU Package Power**: Total CPU power including:
  - All cores
  - L3 cache
  - Memory controller (integrated)
  - Integrated GPU (if present)

- **NOT Included**:
  - Discrete GPU
  - RAM (DDR modules)
  - Storage
  - Display
  - Other system components

## Comparison

| Method | Measures | Accuracy | Requirements |
|--------|----------|----------|--------------|
| **Intel RAPL** | CPU Package Only | High for CPU | Intel CPU, WinRing0, Admin |
| **Power Meter** | Total System | Good for system | Modern laptop/device |

## AMD CPUs

AMD Ryzen CPUs have similar functionality but different MSR addresses. Would need separate implementation. The fallback system power measurement works on AMD systems.
