# PCPow - PC Power Controller (v1.2.3)

A powerful utility for Windows that safely manages PC power states and provides Wake-on-LAN functionality.

## Features

- **Power State Management**: Sleep, restart, or shut down your PC with simple commands
- **Application Safety**: Gracefully closes applications before power actions
- **Wake-on-LAN**: Remotely wake PCs on your network using MAC addresses
- **Multiple Interfaces**: PowerShell commands, command-line interface, and desktop shortcuts
- **Terminal-Friendly**: Commands run without closing your terminal window

## Quick Start Guide

### Installation (PowerShell - Run as Administrator)
```powershell
# Download and run the installer
# Ensure you are in the directory where you downloaded the files
.\Install-PCPow.ps1

# Or specify a custom installation path:
.\Install-PCPow.ps1 -InstallPath "C:\Tools\PCPow"

# Force overwrite if already installed:
.\Install-PCPow.ps1 -Force

# Restart PowerShell after installation
```

### PowerShell Commands (Recommended)
```powershell
# Quick aliases - most convenient way to use PCPow
pows          # Put your PC to sleep
powr          # Restart your PC
powd          # Shut down your PC
poww MAC      # Wake a remote PC (replace MAC with the actual MAC address)

# Full commands with more options
Start-PCSleep [-Force] [-SkipAction]
Restart-PC [-Force] [-SkipAction]
Stop-PC [-Force] [-SkipAction]
Wake-PC -MACAddress "00:11:22:33:44:55"
```

### Command Line Interface (CMD)
```cmd
pcpow sleep [-force] [-skipaction]    # Put PC to sleep
pcpow restart [-force] [-skipaction]  # Restart PC
pcpow shutdown [-force] [-skipaction] # Shut down PC
pcpow wake MAC-ADDRESS                # Wake a remote PC
```

### Command Options
- `-Force`: Skip countdown and force-close applications
- `-SkipAction`: Test mode - shows what would happen without actually performing the action

## Wake-on-LAN Setup

To configure your PC to be woken up remotely:

```powershell
# Run this command in PowerShell
Setup-WakeOnLAN [-AllowRemoteAccess]
```

This configures your network adapters and saves MAC addresses to `wake-targets.txt` in the `config` directory.

## Configuration

Configuration file is located at: `<InstallPath>\config\pcpow.config.json`
Default location: `$HOME\PCPow\config\pcpow.config.json`

You can customize:
- Countdown duration (`CountdownSeconds`)
- Applications to ignore when closing (`IgnoreApps`)
- Default MAC addresses for Wake-on-LAN (`WakeTargets`)

## Troubleshooting

If commands aren't working:

1. **Restart PowerShell**: Ensure the latest profile changes are loaded.
2. **Verify installation**: Run `Get-Command -Name pows, powr, powd, poww` to check if commands are defined.
3. **Check PATH**: Ensure `<InstallPath>\bin` is in your PATH environment variable.
4. **Permissions**: If having issues, try running PowerShell as Administrator.

## Support

For issues or questions, please create an issue on the GitHub repository: [https://github.com/voolyvex/pcpow/issues](https://github.com/voolyvex/pcpow/issues)

## License

[MIT License](LICENSE)
