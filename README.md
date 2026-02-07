# Per-Process CPU Power Meter

A PowerShell-based tool to measure and track power consumption of individual processes on Windows systems.

## Features

- **Real-time monitoring** of process power consumption
- **Interactive process selection** from a list of running processes
- **Accumulated energy tracking** in millijoules, joules, and kilojoules
- **Live statistics** including:
  - Current system power draw
  - Per-process CPU utilization
  - Process power share
  - Total energy consumed
  - Average power consumption
  - Monitoring duration

## Requirements

- Windows system with Power Meter performance counters (typically modern laptops and devices)
- PowerShell 5.1 or later
- Administrator privileges may be required for some systems

## How It Works

The tool uses a simple formula to calculate per-process power consumption:

```
Process Energy = Time × Total Power × (Process CPU / Total CPU)
```

1. **CPU Utilization**: Uses `Get-Counter` to measure per-process CPU usage and total system CPU usage
2. **Power Consumption**: Reads system power draw from Windows Power Meter performance counters
3. **Calculation**: Combines data to estimate process power share based on CPU utilization ratio
4. **Accumulation**: Continuously adds energy consumed over each measurement interval

## Usage

### Basic Usage

Run the script in PowerShell:

```powershell
.\ProcessPowerMeter.ps1
```

### Custom Measurement Interval

Specify a different measurement interval (default is 2 seconds):

```powershell
.\ProcessPowerMeter.ps1 -MeasurementIntervalSeconds 5
```

### Interactive Menu

1. The tool displays a list of running processes sorted by CPU usage
2. Select a process by entering its number
3. Monitoring begins and displays real-time statistics
4. Press 'Q' to stop monitoring and return to the menu
5. Press 'Q' again at the main menu to exit

## Example Output

```
=== Monitoring Process: chrome ===
============================================================
Press 'Q' to stop monitoring and return to menu

Current Measurements:
  System Power:          15.43 W
  Process CPU Usage:     12.45%
  Total CPU Usage:       25.30%
  Process Power Share:   7.59 W

Accumulated Statistics:
  Total Energy Consumed: 45.23 J
  Average Power:         7.54 W
  Monitoring Duration:   00:00:06
  Measurements Taken:    3
```

## Limitations

- Power measurements are system-wide estimates allocated by CPU usage
- Doesn't account for GPU, disk I/O, or network power consumption
- Accuracy depends on the system's Power Meter implementation
- Some systems may not have Power Meter counters available

## Troubleshooting

**"Power Meter counters are not available"**
- Your system may not support Power Meter counters
- Try running PowerShell as Administrator
- Some desktop systems may not have this feature (more common on laptops)

**Inaccurate readings**
- Ensure minimal background activity for more accurate measurements
- Longer measurement intervals provide more stable readings
- Power readings include all system components, not just CPU

## Technical Details

- **Performance Counters Used**:
  - `\Power Meter(_total)\Power` - System power in milliwatts
  - `\Process(*)\% Processor Time` - Per-process CPU utilization

- **Energy Units**:
  - Measurements in millijoules (mJ)
  - Automatically converts to joules (J) and kilojoules (kJ)

## License

Free to use and modify.
