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

## Using RAPL with LibreHardwareMonitor

The current scripts use **LibreHardwareMonitor** for RAPL access, which is simpler than building custom libraries.

1. **Ensure LibreHardwareMonitorLib.dll is present**:
   ```powershell
   # Navigate to your script directory
   cd "<path-to-your-script-directory>"
   
   # Check if DLL exists
   Test-Path .\LibreHardwareMonitorLib.dll
   ```

2. **Download if missing**:
   - Get from: https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases
   - Extract `LibreHardwareMonitorLib.dll` to your script directory

3. **The DLL handles driver installation automatically** when run as Administrator

## Running with RAPL Support

```powershell
# Run as Administrator
.\ProcessPowerMeter-Interactive.ps1
# Or any other script that uses RAPL
```

The scripts will:
- ✓ Auto-detect RAPL support via LibreHardwareMonitor
- ✓ Initialize hardware monitoring and CPU sensors
- ✓ Read CPU package power directly from RAPL registers
- ✓ Fall back gracefully if RAPL unavailable
- ✓ Show clear error messages if hardware/drivers not accessible

## Verification

To check if RAPL is working:

```powershell
# Run the test script as Administrator
.\prototypes\Test-LibreHardwareMonitor.ps1
```

This will:
- Initialize LibreHardwareMonitor
- Detect your CPU
- List all available sensors
- Show current CPU package power reading
- Confirm RAPL is functioning correctly

## Troubleshooting

### "RAPL not supported"
- **Check CPU**: Must be Intel Sandy Bridge or newer
- **Install driver**: Follow WinRing0 installation steps above
- **Run as Admin**: Required for MSR access

### "Failed to initialize RAPL"
- Verify WinRing0 driver is loaded: `sc query WinRing0_1_2_0`
- Try running OpenHardwareMonitor first
- Check Windows Event Viewer for driver errors

### "LibreHardwareMonitorLib.dll not found"
- Download from: https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases
- Place `LibreHardwareMonitorLib.dll` in the same directory as the scripts
- Verify .NET Framework 4.x is installed (usually pre-installed on Windows 10/11)

### Still not working?
- Use the basic script: `.\prototypes\ProcessPowerMeter.ps1`
- It uses system-wide Power Meter counters (no RAPL required)
- Works on any system without special privileges

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
