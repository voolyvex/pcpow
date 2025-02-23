# PCPow

Safely close apps and put your Windows PC to sleep, restart, or shutdown.

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

## Requirements

- Windows 10+
- PowerShell 5.1+
- Admin rights (install only) 