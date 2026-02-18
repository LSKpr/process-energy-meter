# Intel PCM Setup Guide

## What is Intel PCM?

Intel Performance Counter Monitor (Intel PCM) is a comprehensive toolkit for monitoring Intel processor performance and energy metrics. It provides:

- **CPU package power** via RAPL (like LibreHardwareMonitor)
- **DRAM power consumption**
- **Memory bandwidth utilization**
- **Cache hit rates and misses**
- **PCIe bandwidth monitoring**
- **HTTP/Prometheus export** for integration with monitoring systems

## Advantages Over LibreHardwareMonitor

- More detailed component-level power breakdown (CPU, DRAM, uncore separately)
- Cross-platform (Linux, Windows, FreeBSD, macOS)
- HTTP API for remote monitoring
- Grafana integration for visualization
- Additional performance metrics beyond power

## Installation

### Windows

1. **Download Pre-compiled Binaries**:
   - Go to https://ci.appveyor.com/project/opcm/pcm/history
   - Click latest successful build
   - Download artifacts (pcm-windows-*.zip)
   - Extract to a folder (e.g., `C:\pcm`)

2. **Install MSR Driver** (Required):
   - PCM needs the MSR (Model-Specific Register) driver
   - On first run, PCM will attempt to install it automatically
   - Must run as Administrator

### Linux

```bash
# Ubuntu/Debian
sudo apt install pcm

# Fedora/RHEL
sudo yum install pcm

# openSUSE
sudo zypper install pcm

# From source
git clone --recursive https://github.com/intel/pcm
cd pcm
mkdir build && cd build
cmake ..
cmake --build . --parallel
```

## Usage with ProcessPowerMeter-Multi.ps1

### Step 1: Start PCM Sensor Server

**Windows**:
1. Right-click PowerShell and select "Run as Administrator"
2. Navigate to PCM directory:
   ```powershell
   cd C:\pcm
   .\pcm-sensor-server.exe -p 9738
   ```
3. Keep the window open (server runs in foreground on Windows)

**Linux**:
```bash
sudo ./pcm-sensor-server -p 9738
# Or run as daemon:
sudo ./pcm-sensor-server -d -p 9738
```

The server will start and listen on port 9738 (default).

### Step 2: Verify PCM is Running

Open browser and navigate to:
```
http://localhost:9738/
```

You should see JSON output with metrics like:
```json
{
  "package_power": 8.5,
  "dram_power": 2.3,
  "cpu_energy_joules": 12345.67,
  ...
}
```

### Step 3: Run PowerShell Script

```powershell
.\ProcessPowerMeter-Multi.ps1
```

When prompted, select **"Intel PCM (Performance Counter Monitor)"** from the menu.

## Command-Line Options

```bash
# Run on custom port
pcm-sensor-server -p 8080

# Listen on specific interface
pcm-sensor-server -l 192.168.1.10 -p 9738

# Run as daemon (Linux/macOS)
pcm-sensor-server -d -p 9738

# Run with real-time priority (Linux)
pcm-sensor-server -R -p 9738
```

## Available Metrics from PCM

When using PCM with the PowerShell script, you get:

### Primary Metrics (used for per-process allocation):
- **CPU Package Power**: Total CPU power including cores + cache
- **DRAM Power**: Memory power consumption
- **Combined Power**: CPU + DRAM for comprehensive measurement

### Additional Context (not currently used but available):
- Core frequency and utilization
- Memory bandwidth (read/write)
- Cache hit/miss rates
- PCIe traffic
- Thermal headroom

## Troubleshooting

### "PCM sensor server not available on port 9738"

**Solution**:
- Check PCM is running: Open http://localhost:9738/ in browser
- Verify firewall isn't blocking port 9738
- Try restarting PCM: Stop and start pcm-sensor-server

### "Access denied" / "MSR driver error"

**Solution**:
- Run as Administrator (Windows) or with sudo (Linux)
- Install/update MSR driver:
  - Windows: pcm-sensor-server will prompt for driver installation
  - Linux: Load msr module: `sudo modprobe msr`

### PCM shows zeros for power metrics

**Solution**:
- Your CPU may not support RAPL power monitoring
- Check CPU compatibility: Intel Sandy Bridge (2011) or newer required
- Verify in BIOS that power monitoring isn't disabled

### "No power metrics found in PCM response"

**Solution**:
- PCM is running but not exposing power metrics
- Check PCM output format (JSON vs Prometheus)
- Update PCM to latest version
- Try running basic PCM tool first: `.\pcm.exe` to verify RAPL works

## Running as Non-Root on Linux

PCM can run without root privileges using Linux perf_event:

```bash
# One-time setup as root
echo -1 | sudo tee /proc/sys/kernel/perf_event_paranoid

# Run as regular user
export PCM_NO_MSR=1
export PCM_KEEP_NMI_WATCHDOG=1
pcm-sensor-server -p 9738
```

**Note**: Some metrics will be unavailable in non-root mode.

## Comparison: PCM vs LibreHardwareMonitor

| Feature | Intel PCM | LibreHardwareMonitor |
|---------|-----------|---------------------|
| CPU Package Power | ✓ | ✓ |
| DRAM Power | ✓ | ✗ |
| Memory Bandwidth | ✓ | ✗ |
| HTTP API | ✓ | ✗ |
| Grafana Support | ✓ | ✗ |
| Windows Support | ✓ | ✓ |
| Linux Support | ✓ | ✓ |
| macOS Support | ✓ | ✗ |
| Setup Complexity | Medium | Low |
| Per-Process Power | ✗ | ✗ |

**Note**: Neither tool provides per-process power directly - both require CPU utilization ratio allocation (which the PowerShell script handles).

## Integration with Grafana

If you want to visualize power consumption over time:

1. **Install Grafana**:
   ```bash
   # Docker
   docker run -d -p 3000:3000 grafana/grafana
   ```

2. **Add Prometheus datasource** pointing to `http://localhost:9738/`

3. **Import PCM dashboard**:
   - Use dashboard from https://github.com/intel/pcm/tree/master/scripts/grafana

4. **View real-time power metrics** in web UI

## Resources

- Official Repo: https://github.com/intel/pcm
- Documentation: https://github.com/intel/pcm/blob/master/doc/
- Windows Guide: https://github.com/intel/pcm/blob/master/doc/WINDOWS_HOWTO.md
- FAQ: https://github.com/intel/pcm/blob/master/doc/FAQ.md
- Grafana Integration: https://github.com/intel/pcm/blob/master/scripts/grafana/README.md

## License

Intel PCM is open source under BSD-3-Clause license.
