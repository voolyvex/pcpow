# PCPow - Windows Power Management Scripts

Robust PowerShell scripts for gracefully closing applications before sleep, restart, or shutdown on Windows 10/11.

## Features

- 🔒 Safely closes running applications
- 💤 Three power actions: sleep, restart, shutdown
- ⚡ Quick commands from terminal or Run menu
- 🛡️ Preserves system processes
- ⚙️ Force mode for unresponsive apps

## 🆕 Version 2.0 Features
- Centralized configuration via JSON
- Enhanced color-coded status messages
- Configurable timeout settings
- Detailed error reporting and logging
- Consistent behavior across all power actions
- Improved process exclusion list
- PowerShell 5.1+ module structure
- Better error handling and recovery

## Configuration
Create or modify `pcpow.config.json` to customize behavior:
```json
{
    "version": "1.0.0",
    "timeoutMS": 5000,
    "colors": {
        "warning": "Yellow",
        "success": "Green",
        "error": "Red",
        "info": "Cyan",
        "action": "Magenta"
    },
    "excludedProcesses": [
        "explorer", "svchost", "csrss", "smss",
        "wininit", "winlogon", "spoolsv", "lsass"
    ]
}
```

## Requirements

- Windows 10/11
- PowerShell 5.1 or later
- Administrator privileges for setup

## Installation

1. Clone the repository
2. Run PowerShell as Administrator
3. Execute setup script:
```powershell
.\setup-shortcuts.ps1
```

## Usage

### Command Prompt or Run Menu (Win+R):
```batch
pcpow sleep
pcpow restart
pcpow shutdown
```

### PowerShell Quick Commands:
```powershell
pow sleep   # or pows
pow restart # or powr
pow shutdown # or powd
```

Add `-Force` to skip confirmation and force close apps:
```powershell
pow sleep -Force  # Force sleep mode
```

## Files

- `pcpow-common.psm1` - Core PowerShell module
- `pcpow.config.json` - Configuration file
- `Close-And*.ps1` - Power action scripts
- `pcpow.bat` - Command-line interface
- `setup-shortcuts.ps1` - Installation script

## Safety Features

- Confirmation prompts before actions
- Graceful application closing with configurable timeout
- System process protection via exclusion list
- Comprehensive error handling and logging
- Configurable timeouts for graceful exits

## Error Handling

- Detailed error messages with color coding
- Safe fallbacks for missing configuration
- Proper exit codes for automation
- Graceful recovery from application close failures

## License

MIT License - Feel free to modify and distribute

## Contributing

Pull requests welcome! Please ensure your changes:
- Follow PowerShell best practices
- Include proper error handling
- Update documentation as needed
- Maintain backward compatibility 