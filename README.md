# PCPow - Windows Power Management Scripts

Robust PowerShell scripts for gracefully closing applications before sleep, restart, or shutdown on Windows 10/11.

## Features

- 🔒 Safely closes running applications
- 💤 Three power actions: sleep, restart, shutdown
- ⚡ Quick commands from terminal or Run menu
- 🛡️ Preserves system processes
- ⚙️ Force mode for unresponsive apps

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

- `Close-AndSleep.ps1` - Sleep mode script
- `Close-AndRestart.ps1` - Restart script
- `Close-AndShutdown.ps1` - Shutdown script
- `pcpow.bat` - Command-line interface
- `setup-shortcuts.ps1` - Installation script

## Safety Features

- Confirmation prompts before actions
- Graceful application closing
- System process protection
- Error handling for each application
- 5-second timeout for graceful exits

## License

MIT License - Feel free to modify and distribute

## Contributing

Pull requests welcome! Please ensure your changes maintain the focus on safety and reliability. 