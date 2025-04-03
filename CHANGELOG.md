# Changelog

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
