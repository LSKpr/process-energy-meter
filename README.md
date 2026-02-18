# Per-Process CPU Power Meter

A PowerShell-based tool for monitoring power consumption of individual processes on Windows systems with support for Intel RAPL (CPU-only) and system-wide power measurements.

## Features

- **Real-time interactive monitoring**: Auto-refreshing display with continuous background data collection
- **Dual measurement modes**: Intel RAPL (CPU package) + System-wide power
- **Energy accumulation**: Track total energy consumed by each process since monitoring started
- **Multiple views**: Top X processes list or detailed single-process focus view
- **Power history graphs**: ASCII-based visualization of power consumption over time
- **Dynamic configuration**: Adjust measurement interval on the fly (1-60 seconds)
- **System statistics**: Runtime, total energy, current CPU/system power at a glance
- **CPU utilization-based power allocation**: Proportional distribution to processes
- **Color-coded output**: Enhanced readability
- **Admin privilege detection**: Automatic verification

## Versions

### ProcessPowerMeter-Interactive.ps1 (⭐ Recommended)
- **Interactive CLI**: Similar to `top`/`htop` with auto-refreshing display
- **Commands**: `list X`, `focus X`, `interval X`, `help`, `quit`
- **Continuous monitoring**: Energy accumulates from program start
- **Background data collection**: Updates every 2 seconds (configurable)
- **Non-blocking input**: Smooth typing while display refreshes
- **Power graphs**: Visual history for focused processes
- **Requires administrator privileges** for RAPL access
- Best for: Extended monitoring sessions, energy consumption analysis over time

### ProcessPowerMeter-Multi.ps1 (Most Flexible)
- **Multiple power sources**: Choose between LibreHardwareMonitor RAPL, Intel PCM, or Windows Power Meter
- Interactive source selection on startup
- Automatically detects available power measurement backends
- Supports switching between sources without restarting
- **Requires administrator privileges** for RAPL/PCM sources
- Best for systems with Intel PCM installed or when comparing different measurement methods

### ProcessPowerMeter-CPU.ps1 (Most Accurate)
- **Dual measurement**: Intel RAPL (CPU package) + System-wide power side-by-side
- **Requires administrator privileges**
- Requires LibreHardwareMonitorLib.dll and OpenHardwareMonitor service
- Shows both CPU-only and total system power simultaneously
- Best for detailed CPU power analysis

### ProcessPowerMeter.ps1 (Simplest)
- Uses system-wide Power Meter counters only
- No administrator privileges required
- Works on any Windows system with Power Meter support
- Single measurement type (system total)
- Best for quick system-wide power estimates

## Requirements

### For ProcessPowerMeter-Interactive.ps1:
- **Intel CPU with RAPL support** (Sandy Bridge or newer, 2011+)
  - Check with: `wmic cpu get name`
  - i3/i5/i7/i9 from 2011 onwards
- **Windows PowerShell 5.1** or later (pre-installed on Windows 10/11)
- **Administrator privileges required** (for RAPL MSR access)
- **LibreHardwareMonitorLib.dll** (included in repository)
- Windows Power Meter counters (optional, for system-wide measurement)

### For ProcessPowerMeter-CPU.ps1:
- Intel CPU with RAPL support (Sandy Bridge or newer, 2011+)
- Windows PowerShell 5.1 or later
- **Administrator privileges required**
- LibreHardwareMonitorLib.dll (included)

### For Prototype Versions:
See individual script headers in the `prototypes/` folder for specific requirements.
Most require administrator privileges and LibreHardwareMonitorLib.dll.

## How It Works

All versions use the same core formula to calculate per-process power:

```
Process Energy = Time × Power × (Process CPU / Total CPU)
```

### Power Measurement Approach

**CPU-Only (RAPL via LibreHardwareMonitor):**
1. Reads CPU package power directly from Intel MSR registers
2. Allocates CPU power to processes based on their CPU utilization ratio
3. Measures only CPU cores + cache (excludes GPU, display, RAM, etc.)
4. Typical range: 5-50W depending on workload

**System-Wide (Windows Power Meter):**
1. Reads total system power from Windows Performance Counters
2. Allocates total power to processes based on CPU utilization ratio
3. Includes all system components (CPU + GPU + display + RAM + etc.)
4. Typical range: 10-100W depending on system
5. Often unavailable on desktop systems, more common on laptops

### Energy Accumulation

The Interactive version continuously tracks energy consumption:
- Updates every 2 seconds (configurable 1-60s)
- Accumulates energy for each process from program start
- Ranks processes by total accumulated energy, not instantaneous power
- Maintains power history for graphing (last 100 measurements)

## Usage

### Quick Start (Interactive CLI - Recommended)

```powershell
# Right-click PowerShell → Run as Administrator
.\ProcessPowerMeter-Interactive.ps1
```

The script will:
1. Check for administrator privileges and LibreHardwareMonitorLib.dll
2. Initialize hardware monitoring and detect CPU
3. Collect initial baseline data (6 seconds)
4. Start with `list 20` view showing top 20 processes
5. Auto-refresh every 2 seconds while accepting commands

**Available Commands:**
- `list X` - Show top X processes by accumulated energy (e.g., `list 10`)
- `focus X` - Detailed view of process #X with power graph
- `interval X` - Change measurement frequency (1-60 seconds)
- `help` - Show command reference
- `quit` or `exit` - Exit program

### Quick Start (Other Versions)

```powershell
# Run as Administrator
.\ProcessPowerMeter-CPU.ps1       # Dual measurement (RAPL + System)

# Prototype versions (in prototypes folder)
.\prototypes\ProcessPowerMeter-Advanced.ps1  # Memory-weighted allocation
.\prototypes\ProcessPowerMeter-Top5.ps1      # Continuous top 5 monitor
.\prototypes\ProcessPowerMeter-Multi.ps1     # Choose power source
```

### Quick Start (Basic Version)

```powershell
# No admin required
.\ProcessPowerMeter.ps1
```

### Interactive Menu

The Interactive CLI provides commands to:
- List top X processes by accumulated energy consumption
- Focus on individual processes with detailed metrics and power graphs
- Adjust measurement frequency in real-time
- View comprehensive system statistics
- Type 'help' for full command reference

Prototype versions in `prototypes/` folder offer different features:
- Multi-source power measurement selection
- Memory-weighted power allocation algorithms
- Various monitoring approaches for comparison

### Example Output (Interactive CLI)

```
========================================================
   Process Power Monitor - Top 20
========================================================

System Statistics:
  Runtime:              00:15:32
  Measurements:         465
  Measurement Interval: 2s
  Tracked Processes:    147

  Total CPU Usage:      23.5%
  CPU Power Now:        8.45 W
  System Power Now:     N/A
  Total CPU Energy:     7.89 kJ

Top 20 Processes:
#     Process                        CPU Energy   CPU %      Power Now
---------------------------------------------------------------------------
1     chrome.exe                      2.15 kJ    8.50%      1.234 W
2     Code.exe                        1.87 kJ    4.20%      0.876 W
3     firefox.exe                     1.23 kJ    3.10%      0.654 W
4     Discord.exe                     0.98 kJ    2.35%      0.432 W
...

Commands: list X | focus X | interval X | help | quit
Command>
```

## Setup Guide

### For ProcessPowerMeter-Multi.ps1

**Option 1: Intel PCM (Recommended for comprehensive metrics)**

1. **Download Intel PCM**:
   - Windows: Get from https://github.com/intel/pcm/releases or https://ci.appveyor.com/project/opcm/pcm/history
   - Linux: `sudo apt install pcm` or `sudo yum install pcm`

2. **Start PCM sensor server**:
   
   **Windows** (Run PowerShell as Administrator):
   ```powershell
   cd C:\path\to\pcm
   .\pcm-sensor-server.exe -p 9738
   ```
   
   **Linux**:
   ```bash
   sudo ./pcm-sensor-server -p 9738
   ```

3. **Run the script**:
   ```powershell
   .\prototypes\ProcessPowerMeter-Multi.ps1
   # Select "Intel PCM" from the menu
   ```

**Option 2: LibreHardwareMonitor (For RAPL)**

1. **Download LibreHardwareMonitor**:
   - Get from https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases
   - Extract `LibreHardwareMonitorLib.dll` to the script directory

2. **Run the script**:
   ```powershell
   # Right-click PowerShell → Run as Administrator
   .\prototypes\ProcessPowerMeter-Multi.ps1
   # Select "LibreHardwareMonitor (RAPL)" from the menu
   ```

**Option 3: Windows Power Meter (No setup)**

Just run the script - Power Meter is built into Windows.

### For ProcessPowerMeter-Interactive.ps1

1. **Download LibreHardwareMonitor** (if you don't have the DLL):
   - Get from https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases
   - Extract `LibreHardwareMonitorLib.dll` to the script directory
   - **Note**: DLL is already included in this repository

2. **Run as Administrator**:
   ```powershell
   # Right-click PowerShell → Run as Administrator
   cd "<path-to-script-directory>"
   .\ProcessPowerMeter-Interactive.ps1
   ```

3. **Start monitoring**:
   - Script auto-starts with `list 20` view
   - Type commands to change views or settings
   - Energy accumulates continuously until you quit

See [RAPL-Setup-Guide.md](RAPL-Setup-Guide.md) for detailed RAPL setup and troubleshooting.

### For ProcessPowerMeter.ps1

No special setup required - just run the script.

## Troubleshooting

### "RAPL sensor not found"
- Ensure you have an Intel CPU (Sandy Bridge or newer)
- Make sure OpenHardwareMonitor is running
- Run PowerShell as Administrator

### "Not running as Administrator"
- ProcessPowerMeter-CPU.ps1 requires admin privileges
- Right-click PowerShell and select "Run as Administrator"

### "Power Meter counter not available"
- Your system may not support Power Meter counters
- Try on a different device (usually works on laptops)

### "LibreHardwareMonitorLib.dll not found"
- Download from https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases
- Place the DLL in the same directory as the script

## Technical Details

### CPU Package Power (RAPL)
- Measured via Intel RAPL MSR registers (0x611)
- Includes CPU cores and cache power
- Does NOT include GPU, memory controller, or other components
- Typical range: 5-50W depending on workload

### System Total Power
- Measured via Windows Power Meter counters
- Includes ALL system components
- CPU + GPU + Display + RAM + Storage + Peripherals
- Typical range: 10-100W depending on system

### Comparison
CPU-only power is typically 30-60% of total system power on modern laptops.

## Files

### Main Scripts
- `ProcessPowerMeter-Interactive.ps1` - **Interactive CLI** with continuous monitoring (⭐ recommended)
- `ProcessPowerMeter-CPU.ps1` - Dual measurement (RAPL + System side-by-side)

### Libraries & Tools
- `LibreHardwareMonitorLib.dll` - Hardware monitoring library for RAPL access

### Prototypes
- `prototypes/` - Older prototype versions and utilities:
  - `ProcessPowerMeter-Multi.ps1` - Multi-source version (RAPL/PCM/PowerMeter)
  - `ProcessPowerMeter-Advanced.ps1` - Memory-weighted power allocation
  - `ProcessPowerMeter-Top5.ps1` - Continuous top 5 process monitor
  - `ProcessPowerMeter.ps1` - Basic version, system-wide only
  - `Test-LibreHardwareMonitor.ps1` - Diagnostic script to verify RAPL
  - `Install-PCM-Driver.ps1` - Helper for Intel PCM driver installation

### Documentation
- `README.md` - This file
- `RAPL-Setup-Guide.md` - Detailed RAPL setup and troubleshooting
- `Intel-PCM-Setup-Guide.md` - Intel PCM installation guide
- `.gitignore` - Git ignore file

## License

This project is open source and available for educational and research purposes.

## Credits

- Uses [LibreHardwareMonitor](https://github.com/LibreHardwareMonitor/LibreHardwareMonitor) for RAPL access
- Windows Performance Counters for system power measurement
