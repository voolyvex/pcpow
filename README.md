# PCPow (v1.2.2)

**Updated: May 17, 2023**

Safely close Windows applications and manage PC power state (sleep/restart/shutdown) with Wake-on-LAN support.

## Key Improvements in v1.2.2
- **Fixed critical issue**: Sleep/restart/shutdown operations now complete successfully
- **Enhanced power management**: Multiple fallback methods to ensure power actions succeed
- **Better application handling**: More reliable application closing with multi-level termination
- **Expanded application support**: Added support for more known applications
- **Admin privilege awareness**: Detects and warns about missing administrator privileges
- **Blocking process detection**: Identifies processes that might interfere with power operations

## Quick Install & Setup

1.  **Download/Clone:** Get the PCPow files into a directory (e.g., `C:\PCPow`).
2.  **Run Setup Script:** Open PowerShell **as Administrator** in the PCPow directory and run:
    ```powershell
    powershell -ExecutionPolicy Bypass -File .\setup-shortcuts.ps1
    ```
    This copies files to `%USERPROFILE%\AppData\Local\Microsoft\WindowsApps\` and updates your PowerShell profile.
3.  **Restart PowerShell:** Close and reopen any PowerShell windows for the aliases (`pows`, `powr`, `powd`, `poww`) to become available.

## Commands

### Command Line
```
pcpow sleep     # Close apps and sleep PC
pcpow restart   # Close apps and restart PC
pcpow shutdown  # Close apps and shutdown PC
pcpow wake MAC  # Wake a remote PC with MAC address
```

### PowerShell
```powershell
# Aliases:
pows            # Sleep
powr            # Restart
powd            # Shutdown
poww [MAC]      # Wake PC

# Functions:
Start-PCSleep
Restart-PC
Stop-PC
Wake-PC [MAC]
```

### Options
```
-Force          # Skip countdown and force close apps
```

## Configure

Edit the configuration file located at:
`%USERPROFILE%\AppData\Local\Microsoft\WindowsApps\pcpow.config.json`

```json
{
  "version": "1.2.2",      # Current software version
  "timeoutMS": 5000,      # Graceful close wait time (ms)
  "AlwaysForce": false,   # Skip countdown timer
  "NoGraceful": false,    # Skip graceful close attempt
  "colors": {           # Customize console output colors
    "warning": "Yellow",
    "success": "Green",
    "error": "Red",
    "info": "Cyan",
    "action": "Magenta",
    "debug": "DarkGray"
  },
  "excludedProcesses": [ # Lowercase process names to always exclude
    "explorer",
    "powershell",
    "cmd",
    // ... other default system processes ...
  ]
}
```

## Wake-on-LAN Setup

1.  **Run WoL Setup Script:** Open PowerShell **as Administrator** in the PCPow install directory (`%USERPROFILE%\AppData\Local\Microsoft\WindowsApps\`) and run:
    ```powershell
    powershell -ExecutionPolicy Bypass -File .\Setup-WakeOnLAN.ps1 -AllowRemoteAccess
    ```
    *(The `-AllowRemoteAccess` flag is optional but recommended for convenience)*
2. Enable Wake-on-LAN in your BIOS/UEFI settings
3. Use `pcpow wake [MAC]` from another PC to wake this computer

## Requirements

- Windows 10 or 11
- PowerShell 5.1+
- Administrator rights (for `setup-shortcuts.ps1` and `Setup-WakeOnLAN.ps1`)

## Troubleshooting

### PowerShell Aliases Not Working

If the PowerShell aliases don't work after restarting:
```powershell
powershell -ExecutionPolicy Bypass -File "%USERPROFILE%\AppData\Local\Microsoft\WindowsApps\update-profile.ps1"
```
Then restart PowerShell again.

### Wake-on-LAN Not Working

1.  Verify Wake-on-LAN is enabled in your PC's BIOS/UEFI settings.
2.  Ensure the target PC is connected via Ethernet (WoL over Wi-Fi is unreliable).
3.  Confirm the MAC address used in the `pcpow wake [MAC]` command is correct for the target PC's Ethernet adapter.
4.  Check if firewall rules created by `Setup-WakeOnLAN.ps1` are active.

### Power Actions Not Working

If sleep, restart, or shutdown actions don't complete:
1. Run PCPow with `-Force` flag to bypass graceful application closing
2. Make sure you're running with administrative privileges for certain power actions
3. Check for any interfering applications listed in the log outputs
4. Review logs at `%USERPROFILE%\AppData\Local\PCPow\logs\` for detailed error information
5. Ensure the system doesn't have pending Windows Updates requiring restart

## File Locations

- **Installation Directory:** `%USERPROFILE%\AppData\Local\Microsoft\WindowsApps\`
  - Contains: `pcpow.bat`, `PCPow.ps1`, `pcpow.config.json`, `Setup-WakeOnLAN.ps1`, `update-profile.ps1`
- **PowerShell Profile:** `%USERPROFILE%\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1`
- **Log Files:** `%USERPROFILE%\AppData\Local\PCPow\logs\`