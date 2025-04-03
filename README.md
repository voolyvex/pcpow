# PCPow (v1.1.0)

Safely close apps and put your Windows PC to sleep, restart, or shutdown.

## What's New in v1.1.0

- **Better Terminal Detection**: Properly handles all PowerShell and Terminal windows
- **Admin Window Support**: Special handling for elevated PowerShell windows 
- **Process Tree Protection**: Improved detection to prevent closing the current process tree
- **Enhanced Sleep Logic**: Multiple sleep methods for more reliable sleep behavior
- **Suspended Process Handling**: Detects and cleans up unresponsive applications

## Install

1. Run as administrator:
   ```powershell
   powershell -ExecutionPolicy Bypass -File setup-shortcuts.ps1
   ```
2. Restart PowerShell

## Use

From Command Prompt or Run (Win+R):
```
pcpow sleep     # Sleep PC
pcpow restart   # Restart PC
pcpow shutdown  # Shutdown PC
```

From PowerShell:
```powershell
# Quick aliases:
pows    # Sleep
powr    # Restart
powd    # Shutdown

# Full commands:
Start-PCSleep
Restart-PC
Stop-PC
```

Add `-Force` to skip confirmations:
```
pcpow sleep -Force
# or
Start-PCSleep -Force
```

## Configure

Edit `pcpow.config.json`:
```json
{
  "timeoutMS": 5000,      # Wait time for apps (ms)
  "AlwaysForce": false,   # Skip confirmations
  "NoGraceful": false     # Skip graceful closing
}
```

## Troubleshoot

- Commands not found? Restart PowerShell
- Permission error? Run as administrator
- Apps won't close? Use `-Force`
- Terminal windows still open? Update to v1.1.0+
- Sleep issues? Use `-Force` to enable aggressive sleep methods

## Requirements

- Windows 10+
- PowerShell 5.1+
- Admin rights (install only)

## Common Issues Fixed

- **Terminal Windows**: Fixed issues with closing PowerShell/Terminal windows
- **Admin Privileges**: Better handling of elevated windows
- **Sleep Reliability**: Multiple fallback methods if primary sleep fails
- **Process Tree Detection**: Now correctly identifies all parent/child processes