# Changelog

## v1.2.2

Released: May 17, 2023

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

## v1.2.1

Released: April 4, 2025

### Fixed
- Critical bug in PowerShell profile causing parameter passing errors
- Replaced fix-profile-v2.ps1 with enhanced update-profile.ps1 utility to thoroughly clean profile
- Eliminated references to obsolete Close-And*.ps1 scripts
- Improved handling of obsolete file references in PowerShell profile
- Switched from array-based parameter passing to proper hashtable splatting
- Improved path handling for script locations
- Better error reporting when files are not found
- Simplified setup process to automatically update profile

## v1.2.0

Released: April 3, 2025

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

## v1.1.0

Released: April 3, 2025

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

## v1.0.0

Released: March 15, 2025

### Added
- Initial release
- Basic power management functionality (sleep, restart, shutdown)
- Configuration system
- Command-line interface
- PowerShell module with aliases
