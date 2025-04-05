# Changelog

## [1.2.4] - 2025-04-06

### Fixed
- Fixed PCPow.ps1 string interpolation error that occurred when displaying error messages
- Fixed pcpow.bat to use the correct paths from installation directory
- Improved batch file command handling for better performance

### Changed
- Removed deprecated setup-shortcuts.ps1 and update-profile.ps1 scripts
- Consolidated all installation functionality into Install-PCPow.ps1
- Enhanced error messages for improved debugging
- Simplified batch file implementation

## [1.2.3] - 2025-04-05

### Fixed
- Critical bug in PowerShell profile causing parameter passing errors
- Replaced fix-profile-v2.ps1 with enhanced update-profile.ps1 utility to thoroughly clean profile
- Eliminated references to obsolete Close-And*.ps1 scripts
- Improved handling of obsolete file references in PowerShell profile
- Switched from array-based parameter passing to proper hashtable splatting
- Improved path handling for script locations
- Better error reporting when files are not found
- Fixed issue with PCPow-Launcher.ps1 string interpolation that caused unexpected token errors
- Fixed PowerShell profile corruption issues
- Improved error handling in all scripts
- Fixed commands closing terminal window
- Verified pows and powr commands working properly
- Corrected pcpow.bat argument handling
- Fixed PCPow-Launcher.ps1 missing Start-Process command

### Added
- Comprehensive installation script (`Install-PCPow.ps1`)
- Better logging and error messages
- SkipAction parameter for testing without executing actual power commands
- Improved documentation (README, CHANGELOG, Project Summary, LinkedIn Post)
- Added `Setup-WakeOnLAN` function to `PCPow.ps1` and PowerShell profile

### Changed
- Relocated main scripts to dedicated bin directory (within the install path)
- Simplified command structure
- Enhanced error reporting
- Improved Wake-on-LAN configuration
- Centralized functions into `Install-PCPow.ps1` for setup

## [1.2.2] - May 17, 2023

### Fixed
- Critical issue with sleep/restart/shutdown actions not completing after closing applications
- Enhanced power management API calls with better error handling and status reporting
- Added administrator privilege detection and warning for power actions
- Improved process termination with multi-level fallback mechanisms
- Added detection and handling of potential blocking applications
- Implemented more reliable Windows API calls for power state transitions
- Better synchronization between application closing and power actions
- Added system preparatory steps before power actions to clear buffers and release resources
- Expanded list of known applications to ensure proper closure
- More aggressive process termination for stubborn applications that resist standard closure methods

## [1.2.1] - April 4, 2025

### Fixed
- Critical bug in PowerShell profile causing parameter passing errors
- Replaced fix-profile-v2.ps1 with enhanced update-profile.ps1 utility to thoroughly clean profile
- Eliminated references to obsolete Close-And*.ps1 scripts
- Improved handling of obsolete file references in PowerShell profile
- Switched from array-based parameter passing to proper hashtable splatting
- Improved path handling for script locations
- Better error reporting when files are not found
- Simplified setup process to automatically update profile

## v1.2.0 - April 3, 2025

### Added
- Wake-on-LAN functionality to wake remote PCs
- Simplified standalone script architecture
- BIOS/UEFI configuration for optimal Wake-on-LAN support
- 5-second countdown timer instead of manual confirmation
- Better support for Windows 10 and 11
- Special handling for FREIA remote access

### Changed
- Removed dependency on PowerShell modules for better reliability
- Combined all power actions into a single unified script
- Improved excluded process handling
- Standardized MAC address formatting for Wake-on-LAN
- More reliable network adapter power management settings
- Comprehensive cleanup of legacy files

### Fixed
- Issues with restart/sleep/shutdown actions failing
- Redundant code across multiple scripts
- Unreliable PowerShell module loading
- Inefficient process tree handling
- Fast Startup interference with Wake-on-LAN

## v1.1.0 - April 3, 2025

### Added
- Enhanced terminal process detection and handling
- Special handling for elevated (administrator) PowerShell windows
- Process tree protection up to great-grandparent processes
- Multiple sleep methods for more reliable sleep behavior
- Suspended process detection and cleanup

### Fixed
- Issue with terminal windows preventing sleep
- Variable name conflict with PowerShell $Host automatic variable
- Improved shutdown/restart/sleep reliability
- Better error handling and logging

### Changed
- Updated configuration with more excluded system processes
- Improved command-line feedback with detailed status messages
- Centralized process tree management
- Version bump to 1.1.0

## v1.0.0 - March 15, 2025

### Added
- Initial release
- Basic power management functionality (sleep, restart, shutdown)
- Configuration system
- Command-line interface
- PowerShell module with aliases
