# PCPow Project Summary

## Overview

PCPow is a Windows power management utility designed to solve the terminal closing issue when executing power commands. It provides a simple, consistent interface for controlling PC power states (sleep, restart, shutdown) while maintaining terminal sessions and gracefully closing applications.

## Project Components

### Core Files
- **PCPow.ps1**: Main script implementing all power control functionality
- **PCPow-Launcher.ps1**: Launcher that prevents terminal windows from closing
- **pcpow.bat**: Batch file for command-line interface
- **pcpow.config.json**: Configuration file for customizing behavior

### Installation and Setup
- **Install-PCPow.ps1**: One-click installer with customization options
- **PowerShell Profile Integration**: Adds aliases and functions to user profile
- **PATH Environment Variable**: Ensures commands are available system-wide

### Documentation
- **README.md**: User documentation with installation and usage instructions
- **CHANGELOG.md**: Version history and release notes
- **LinkedIn-Post.md**: Marketing content for sharing the project

## Key Features

1. **Terminal-Friendly Commands**
   - Execute power commands without losing terminal sessions
   - PowerShell aliases: pows, powr, powd, poww

2. **Multiple Interfaces**
   - PowerShell functions and aliases
   - Command prompt batch commands
   - Desktop shortcuts (optional)

3. **Graceful Application Handling**
   - Safely closes applications before power actions
   - Configurable countdown timer
   - Force option for bypassing countdown

4. **Wake-on-LAN Support**
   - Wake remote computers using MAC addresses
   - Setup utility for configuring network adapters
   - MAC address management for multiple targets

5. **Robust Error Handling**
   - Comprehensive logging
   - Informative error messages
   - Prevention of PowerShell profile corruption

## Technical Improvements

### Fixed Issues
- String interpolation errors in PCPow-Launcher.ps1
- PowerShell profile corruption
- Terminal closing during power operations
- Path handling and environment variable conflicts

### Added Capabilities
- Test mode with -SkipAction parameter
- Enhanced error handling and reporting
- Improved documentation and examples
- One-click installer with customization options

## User Benefits

- **Preserved Workflow**: Maintain terminal sessions during power operations
- **Simplified Commands**: Easy-to-remember aliases for common actions
- **Cross-Platform Management**: Manage power states for multiple systems
- **Application Safety**: Prevent data loss from forced application closures
- **Customization**: Configure countdown times and application handling

## Future Enhancements

- GUI interface for power management
- Remote management capabilities for network administrators
- Integration with Windows scheduled tasks
- Advanced logging and reporting features
- Multi-language support

## Installation and Usage

### Installation
```powershell
# Default installation
.\Install-PCPow.ps1

# Custom installation path
.\Install-PCPow.ps1 -InstallPath "D:\Tools\PCPow"

# Force reinstallation
.\Install-PCPow.ps1 -Force
```

### Basic Commands
```powershell
# PowerShell
pows           # Sleep
powr           # Restart
powd           # Shutdown
poww <MAC>     # Wake remote PC

# Command Prompt
pcpow sleep    # Sleep
pcpow restart  # Restart
pcpow shutdown # Shutdown
pcpow wake MAC # Wake remote PC
```

### Advanced Usage
```powershell
# Skip countdown and force close applications
pows -Force

# Test mode (doesn't execute power action)
pows -SkipAction

# Configure Wake-on-LAN
Setup-WakeOnLAN -AllowRemoteAccess
```

## Conclusion

PCPow represents a significant improvement in Windows power management, addressing the long-standing issue of terminal sessions closing during power operations. With its simple interface, robust error handling, and flexible deployment options, PCPow enhances productivity for developers, system administrators, and regular users who rely on terminal sessions for their daily work. 