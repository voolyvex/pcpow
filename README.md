# PCPow - Windows Power Management

## Safely sleep/restart/shutdown Windows with open app cleanup

**Key Features**:
- Closes user applications before power actions
- Preserves system processes and services
- Configurable timeouts and process whitelist
- Works from PowerShell or Run dialog (Win+R)

## Quick Install
```powershell
# Run in Admin PowerShell
irm https://raw.githubusercontent.com/voolyvex/pcpow/main/setup-shortcuts.ps1 -OutFile setup.ps1
.\setup.ps1
```

## Basic Usage
```bash
# Command Prompt/Run dialog:
pcpow sleep       # Close apps and sleep
pcpow restart -F  # Force restart without confirmation

# PowerShell:
pows              # Alias for Sleep-PC
powr -Force       # Force restart apps
Stop-PCApps       # Full shutdown command
```

## Configuration (optional)
Create `pcpow.config.json` to customize:
```json
{
    "timeoutMS": 5000,
    "excludedProcesses": [
        "explorer", "backgroundsvc",
        "securityprocess", "antivirus"
    ]
}
```

## Safety
- Confirms destructive actions unless using `-Force`
- Gives apps 5s to close gracefully (configurable)
- Protects critical system processes
- Logs errors to Windows Event Viewer

## License
MIT - Free for personal/professional use 