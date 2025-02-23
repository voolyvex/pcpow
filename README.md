# PCPow

Simple Windows power management tool that safely closes apps before sleep/restart/shutdown.

## Features

- Safely closes all applications before power actions
- Graceful handling of File Explorer windows
- Configurable timeouts and force modes
- Works from PowerShell, Command Prompt, or Run dialog
- Supports sleep, restart, and shutdown operations

## Requirements

- Windows 10 or later
- PowerShell 5.1 or later
- Administrator privileges (for installation only)

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/pcpow.git
   cd pcpow
   ```

2. Run the setup script as administrator:
   ```powershell
   powershell -ExecutionPolicy Bypass -File setup-shortcuts.ps1
   ```

3. Restart PowerShell to load the new commands

## Quick Start

Use any of these commands:

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
powr -Force
```

## Configuration

Edit `pcpow.config.json` to customize behavior:
```json
{
  "version": "1.0.0",
  "timeoutMS": 5000,      # Wait time for apps to close (ms)
  "AlwaysForce": false,   # Skip confirmations always
  "NoGraceful": false,    # Skip graceful app closing
  "colors": {
    "warning": "Yellow",
    "success": "Green",
    "error": "Red",
    "info": "Cyan",
    "action": "Magenta"
  },
  "excludedProcesses": [  # System processes to ignore
    "svchost",
    "csrss",
    "smss",
    "wininit",
    "winlogon"
  ]
}
```

## Configuration Locations

The tool looks for configuration in these locations:
1. Script directory: `pcpow.config.json`
2. User directory: `%USERPROFILE%\AppData\Local\Microsoft\WindowsApps\pcpow.config.json`

## Troubleshooting

1. If commands aren't recognized:
   - Restart PowerShell
   - Verify installation path is in PATH
   - Run setup script again

2. If permission errors occur:
   - Run the affected command as administrator
   - Check file permissions in installation directory

3. If apps don't close properly:
   - Increase `timeoutMS` in config
   - Use `-Force` option
   - Enable `NoGraceful` in config

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

MIT License - See LICENSE file for details

## Security

- Never stores sensitive information
- Runs with user privileges (except installation)
- Configurable process exclusion list
- Graceful application closing by default 