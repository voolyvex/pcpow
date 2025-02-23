# PCPow

Simple Windows power management tool that safely closes apps before sleep/restart/shutdown.

## Quick Start

1. Run `setup-shortcuts.ps1` as administrator
2. Restart PowerShell
3. Use any of these commands:

From Command Prompt or Run (Win+R):
```
pcpow sleep     # Sleep PC
pcpow restart   # Restart PC
pcpow shutdown  # Shutdown PC
```

From PowerShell:
```powershell
pows    # Sleep
powr    # Restart
powd    # Shutdown
```

## Options

Add `-Force` to skip confirmations:
```
pcpow sleep -Force
```

## Configuration

Edit `pcpow.config.json` to customize:
```json
{
  "AlwaysForce": false,  # Skip confirmations always
  "NoGraceful": false,   # Skip graceful app closing
  "timeoutMS": 5000     # Wait time for apps (ms)
}
```

## License

MIT License 