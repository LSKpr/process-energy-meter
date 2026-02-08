# Per-Process CPU Power Meter

A PowerShell-based tool for monitoring power consumption of individual processes on Windows systems with support for both Intel RAPL (CPU-only) and system-wide power measurements.

## Features

- **Dual measurement modes**: Intel RAPL (CPU package only) + System-wide power
- Real-time power consumption monitoring per process
- CPU utilization-based power allocation
- Interactive process selection menu
- Accumulated energy consumption tracking (separate for CPU and system)
- Average power calculations
- Color-coded output for better readability
- Admin privilege detection and verification

## Versions

### ProcessPowerMeter-CPU.ps1 (Recommended)
- **Dual measurement**: Intel RAPL (CPU package) + System-wide power
- **Requires administrator privileges**
- Requires LibreHardwareMonitorLib.dll and OpenHardwareMonitor service
- Shows both CPU-only and total system power side-by-side
- More accurate and comprehensive measurements

### ProcessPowerMeter.ps1 (Basic)
- Uses system-wide Power Meter counters only
- No administrator privileges required
- Works on any Windows system with Power Meter support
- Single measurement type (system total)

## Requirements

### For ProcessPowerMeter-CPU.ps1:
- Intel CPU with RAPL support (Sandy Bridge or newer, 2011+)
- Windows with Power Meter performance counters
- PowerShell 5.1 or later
- **Administrator privileges required**
- LibreHardwareMonitorLib.dll (included)
- OpenHardwareMonitor running as background service

### For ProcessPowerMeter.ps1:
- Windows system with Power Meter performance counters
- PowerShell 5.1 or later
- No special privileges required

## How It Works

Both versions use the same core formula to calculate per-process power:

```
Process Energy = Time × Power × (Process CPU / Total CPU)
```

### ProcessPowerMeter-CPU.ps1 (Dual Measurement)

For RAPL Guide, please view the md file: [RAPL Guide](CPU/RAPL-Setup-Guide.md)

**CPU-Only (RAPL):**
1. Reads CPU package power directly from Intel MSR registers via LibreHardwareMonitor
2. Allocates CPU power to processes based on their CPU utilization ratio
3. Measures only CPU cores + cache (excludes GPU, display, RAM, etc.)

**System-Wide:**
1. Reads total system power from Windows Power Meter counters
2. Allocates total power to processes based on CPU utilization ratio
3. Includes all system components (CPU + GPU + display + RAM + etc.)

Both measurements run simultaneously for direct comparison.

### ProcessPowerMeter.ps1 (Basic)

1. **CPU Utilization**: Uses `Get-Counter` to measure per-process CPU usage
2. **Power Consumption**: Reads total system power from Windows Power Meter
3. **Calculation**: Allocates power proportionally based on CPU utilization
4. **Accumulation**: Tracks energy consumed over time

## Usage

### Quick Start (CPU Version)

```powershell
# Run as Administrator
.\ProcessPowerMeter-CPU.ps1
```

The script will:
1. Check for administrator privileges
2. Initialize LibreHardwareMonitor library
3. Verify RAPL sensor availability
4. Display process selection menu
5. Monitor selected process with dual measurements

### Quick Start (Basic Version)

```powershell
# No admin required
.\ProcessPowerMeter.ps1
```

### Interactive Menu

Both versions provide an interactive menu to:
- Select a process from the list of running processes
- View current power consumption in real-time
- See accumulated energy statistics
- Press 'Q' to stop monitoring and view summary

### Example Output (CPU Version)

```
=== Per-Process Power Monitor ===
Process: Discord

Mode: Dual Measurement (CPU Package + System Total)
Press 'Q' to stop monitoring and return to menu

Current Measurements:
  CPU Package Power (RAPL): 5.26 W
  System Total Power:       13.42 W
  Process CPU Usage:        2.15%
  Total CPU Usage:          8.73%
  Process CPU Power:        0.13 W
  Process System Power:     0.33 W

Accumulated Statistics (CPU Only - RAPL):
  Total Energy Consumed:    1.25 kJ
  Average Power:            0.15 W

Accumulated Statistics (System Total):
  Total Energy Consumed:    3.18 kJ
  Average Power:            0.38 W

General Statistics:
  Monitoring Duration:      02:18:45
  Measurements Taken:       4113
```

## Setup Guide

### For ProcessPowerMeter-CPU.ps1

1. **Download LibreHardwareMonitor**:
   - Get from https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases
   - Extract `LibreHardwareMonitorLib.dll` to the script directory

2. **Install OpenHardwareMonitor Service**:
   ```powershell
   # Download and run OpenHardwareMonitor
   # Keep it running in background for driver access
   ```

3. **Run as Administrator**:
   ```powershell
   # Right-click PowerShell → Run as Administrator
   .\ProcessPowerMeter-CPU.ps1
   ```

See [RAPL-Setup-Guide.md](RAPL-Setup-Guide.md) for detailed setup instructions.

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

- `ProcessPowerMeter-CPU.ps1` - Advanced version with dual measurements
- `ProcessPowerMeter.ps1` - Basic version, system-wide only
- `LibreHardwareMonitorLib.dll` - Hardware monitoring library
- `Test-LibreHardwareMonitor.ps1` - Diagnostic script to verify RAPL
- `RAPL-Setup-Guide.md` - Detailed setup instructions
- `.gitignore` - Git ignore file

## License

This project is open source and available for educational and research purposes.

## Credits

- Uses [LibreHardwareMonitor](https://github.com/LibreHardwareMonitor/LibreHardwareMonitor) for RAPL access
- Windows Performance Counters for system power measurement
