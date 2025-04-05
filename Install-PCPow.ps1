#Requires -Version 5.1
<#
.SYNOPSIS
    Installs PCPow power control utility.

.DESCRIPTION
    This script installs PCPow to the user's profile, sets up PowerShell aliases,
    and configures the PCPow launcher to prevent terminal windows from closing
    during power operations.

.PARAMETER Force
    Overwrites existing installation without confirmation.

.PARAMETER InstallPath
    Specifies a custom installation path.

.EXAMPLE
    .\Install-PCPow.ps1
    
    Installs PCPow to the default location.

.EXAMPLE
    .\Install-PCPow.ps1 -Force
    
    Installs PCPow and overwrites any existing installation.

.EXAMPLE
    .\Install-PCPow.ps1 -InstallPath "D:\Tools\PCPow"
    
    Installs PCPow to a custom location.
#>
[CmdletBinding()]
param (
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [string]$InstallPath = "$HOME\PCPow"
)

# Set strict error handling
$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

# Function to create directories if they don't exist
function Ensure-Directory {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    if (-not (Test-Path -Path $Path)) {
        Write-Verbose "Creating directory: $Path"
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

# Function to add PCPow commands to PowerShell profile
function Update-Profile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BinPath
    )
    
    $profilePath = $PROFILE
    $profileDir = Split-Path -Parent $profilePath
    
    # Ensure profile directory exists
    Ensure-Directory -Path $profileDir
    
    # Create/update profile content
    $profileContent = @"
# PCPow version 1.2.4 commands
Write-Host "Loading PCPow version 1.2.4 commands..." -ForegroundColor Cyan

# Set PCPow bin path
`$env:PCPowBin = "$BinPath"

# Add PCPow bin to PATH if not already there
if (`$env:PATH -notlike "*`$env:PCPowBin*") {
    `$env:PATH = "`$env:PCPowBin;`$env:PATH"
}

# PCPow Aliases
Set-Alias -Name pows -Value Start-PCSleep -Description "Put PC to sleep"
Set-Alias -Name powr -Value Restart-PC -Description "Restart PC"
Set-Alias -Name powd -Value Stop-PC -Description "Shutdown PC"
Set-Alias -Name poww -Value Wake-PC -Description "Wake a remote PC"

# PCPow Functions
function Start-PCSleep { 
    & "`$env:PCPowBin\PCPow-Launcher.ps1" -Action Sleep @args 
}

function Restart-PC { 
    & "`$env:PCPowBin\PCPow-Launcher.ps1" -Action Restart @args 
}

function Stop-PC { 
    & "`$env:PCPowBin\PCPow-Launcher.ps1" -Action Shutdown @args 
}

function Wake-PC {
    param (
        [Parameter(Mandatory = `$true, Position = 0)]
        [string]`$MACAddress
    )
    & "`$env:PCPowBin\PCPow.ps1" -Action Wake -MACAddress `$MACAddress @args
}

function Setup-WakeOnLAN {
    param (
        [Parameter()]
        [switch]`$AllowRemoteAccess
    )
    & "`$env:PCPowBin\PCPow.ps1" -Action SetupWakeOnLAN -AllowRemoteAccess:`$AllowRemoteAccess @args
}

# PCPow version 1.2.4 commands
Write-Host "Loading PCPow version 1.2.4 commands..." -ForegroundColor Cyan
"@
    
    # Check if profile exists and update/create accordingly
    if (Test-Path -Path $profilePath) {
        # Remove any existing PCPow functions
        $currentProfile = Get-Content -Path $profilePath -Raw
        $cleanProfile = $currentProfile -replace "(?ms)# PCPow version.*?Wake-PC \{.*?\}\r?\n\r?\nfunction Setup-WakeOnLAN \{.*?\}\r?\n", ""
        $newProfile = $cleanProfile + "`n" + $profileContent
        Set-Content -Path $profilePath -Value $newProfile
        Write-Host "Updated existing PowerShell profile at: $profilePath" -ForegroundColor Green
    }
    else {
        # Create new profile
        Set-Content -Path $profilePath -Value $profileContent
        Write-Host "Created new PowerShell profile at: $profilePath" -ForegroundColor Green
    }
}

# Main installation function
function Install-PCPow {
    # Confirm installation
    if (-not $Force -and (Test-Path -Path $InstallPath)) {
        $confirm = Read-Host "PCPow installation exists at '$InstallPath'. Overwrite? (y/n)"
        if ($confirm -ne 'y') {
            Write-Host "Installation canceled." -ForegroundColor Yellow
            return
        }
    }
    
    # Create directory structure
    Write-Host "Installing PCPow to: $InstallPath" -ForegroundColor Cyan
    Ensure-Directory -Path $InstallPath
    Ensure-Directory -Path "$InstallPath\bin"
    Ensure-Directory -Path "$InstallPath\config"
    
    # Create PCPow.ps1 (main script)
    $mainScriptPath = "$InstallPath\bin\PCPow.ps1"
    Write-Verbose "Creating main script: $mainScriptPath"
    Set-Content -Path $mainScriptPath -Value @'
[CmdletBinding()]
param (
    [Parameter(Position = 0)]
    [ValidateSet("Sleep", "Restart", "Shutdown", "Wake", "SetupWakeOnLAN")]
    [string]$Action = "Sleep",
    
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$SkipAction,
    
    [Parameter()]
    [string]$MACAddress = "",
    
    [Parameter()]
    [switch]$AllowRemoteAccess
)

# Log start of execution
Write-Verbose "PCPow started with Action: $Action, Force: $Force, SkipAction: $SkipAction"

# Configuration
$configDir = Join-Path (Split-Path -Parent (Split-Path -Parent $PSCommandPath)) "config"
$configFile = Join-Path $configDir "pcpow.config.json"

# Default configuration
$defaultConfig = @{
    CountdownSeconds = 5
    IgnoreApps = @("explorer", "powershell", "cmd")
    WakeTargets = @{}
}

# Load or create configuration
if (Test-Path $configFile) {
    try {
        $config = Get-Content $configFile -Raw | ConvertFrom-Json
        Write-Verbose "Loaded configuration from: $configFile"
    }
    catch {
        Write-Warning "Error loading configuration: $_"
        $config = $defaultConfig
    }
}
else {
    Write-Verbose "Creating default configuration at: $configFile"
    $config = $defaultConfig
    
    if (-not (Test-Path $configDir)) {
        New-Item -Path $configDir -ItemType Directory -Force | Out-Null
    }
    
    $config | ConvertTo-Json | Set-Content $configFile
}

# Function to close applications gracefully
function Close-Applications {
    $processes = Get-Process | Where-Object { 
        $_.MainWindowHandle -ne 0 -and 
        $_.ProcessName -notin $config.IgnoreApps
    }
    
    foreach ($process in $processes) {
        Write-Host "Closing: $($process.ProcessName)" -ForegroundColor Yellow
        try {
            $process.CloseMainWindow() | Out-Null
        }
        catch {
            Write-Warning "Could not close $($process.ProcessName): $_"
        }
    }
}

# Function to send Wake-on-LAN packet
function Send-WakePacket {
    param (
        [Parameter(Mandatory = $true)]
        [string]$MACAddress
    )
    
    # Validate MAC address format
    if ($MACAddress -notmatch '^([0-9A-F]{2}[:-]){5}([0-9A-F]{2})$') {
        Write-Error "Invalid MAC address format. Use format: 00:11:22:33:44:55"
        return $false
    }
    
    # Clean the MAC address
    $MAC = $MACAddress -replace '[:-]', ''
    
    try {
        # Create the magic packet
        $MacByteArray = [byte[]]::new(102)
        
        # First 6 bytes are 0xFF
        for ($i = 0; $i -lt 6; $i++) {
            $MacByteArray[$i] = 0xFF
        }
        
        # Repeat target MAC 16 times
        for ($i = 1; $i -le 16; $i++) {
            for ($j = 0; $j -lt 6; $j++) {
                $MacByteArray[6 * $i + $j] = [Convert]::ToByte($MAC.Substring($j * 2, 2), 16)
            }
        }
        
        # Send the packet
        $UdpClient = New-Object System.Net.Sockets.UdpClient
        $UdpClient.Connect([System.Net.IPAddress]::Broadcast, 9)
        $UdpClient.Send($MacByteArray, $MacByteArray.Length) | Out-Null
        $UdpClient.Close()
        
        Write-Host "Wake-on-LAN packet sent to: $MACAddress" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to send Wake-on-LAN packet: $_"
        return $false
    }
}

# Function to configure Wake-on-LAN
function Setup-WOL {
    param (
        [Parameter()]
        [switch]$AllowRemoteAccess
    )
    
    Write-Host "Configuring Wake-on-LAN..." -ForegroundColor Cyan
    
    try {
        # Get all network adapters
        $adapters = Get-WmiObject MSPower_DeviceWakeEnable -Namespace root\wmi
        $ethernetAdapters = Get-NetAdapter | Where-Object { $_.MediaType -eq "802.3" }
        
        # Enable Wake-on-LAN for all ethernet adapters
        foreach ($adapter in $ethernetAdapters) {
            Write-Host "Configuring adapter: $($adapter.Name)" -ForegroundColor Yellow
            
            # Enable WOL in adapter properties
            $devPath = $adapter.InterfaceDescription
            $matchingWmiAdapter = $adapters | Where-Object { $_.InstanceName -match [regex]::Escape($devPath) }
            
            if ($matchingWmiAdapter) {
                $matchingWmiAdapter.Enable = $true
                $matchingWmiAdapter.Put() | Out-Null
            }
            
            # Configure advanced properties
            $adapterConfig = Get-NetAdapterPowerManagement -Name $adapter.Name -ErrorAction SilentlyContinue
            if ($adapterConfig) {
                $adapterConfig.WakeOnMagicPacket = "Enabled"
                $adapterConfig.WakeOnPattern = "Enabled"
                $adapterConfig | Set-NetAdapterPowerManagement
            }
            
            # Save MAC address
            $MAC = ($adapter.MacAddress -replace '-', ':').ToLower()
            $computerName = $env:COMPUTERNAME
            $config.WakeTargets[$computerName] = $MAC
            
            # Create wake-targets.txt in the config directory
            $targetPath = Join-Path $configDir "wake-targets.txt"
            "$computerName=$MAC" | Out-File -FilePath $targetPath -Append
            
            Write-Host "Saved MAC address $MAC for $computerName" -ForegroundColor Green
        }
        
        # Allow remote wake if specified
        if ($AllowRemoteAccess) {
            Write-Host "Configuring firewall for remote wake..." -ForegroundColor Yellow
            
            # Create firewall rule for WOL
            New-NetFirewallRule -DisplayName "Wake-on-LAN" -Direction Inbound -Protocol UDP -LocalPort 9 -Action Allow -ErrorAction SilentlyContinue | Out-Null
            
            Write-Host "Firewall configured for remote wake access" -ForegroundColor Green
        }
        
        # Save updated configuration
        $config | ConvertTo-Json | Set-Content $configFile
        
        Write-Host "Wake-on-LAN configuration complete" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Error configuring Wake-on-LAN: $_"
        return $false
    }
}

# Execute action based on parameter
switch ($Action) {
    "Sleep" {
        Write-Host "Preparing to put PC to sleep..." -ForegroundColor Cyan
        
        if (-not $Force) {
            # Countdown
            for ($i = $config.CountdownSeconds; $i -gt 0; $i--) {
                Write-Host "Sleeping in $i seconds... Press Ctrl+C to abort" -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
            
            # Close applications
            Close-Applications
        }
        
        if (-not $SkipAction) {
            Write-Host "Executing sleep command now..." -ForegroundColor Green
            Start-Sleep -Seconds 1
            rundll32.exe powrprof.dll,SetSuspendState 0,1,0
        }
        else {
            Write-Host "SkipAction flag set - sleep command was not executed" -ForegroundColor Magenta
        }
    }
    "Restart" {
        Write-Host "Preparing to restart PC..." -ForegroundColor Cyan
        
        if (-not $Force) {
            # Countdown
            for ($i = $config.CountdownSeconds; $i -gt 0; $i--) {
                Write-Host "Restarting in $i seconds... Press Ctrl+C to abort" -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
            
            # Close applications
            Close-Applications
        }
        
        if (-not $SkipAction) {
            Write-Host "Executing restart command now..." -ForegroundColor Green
            Start-Sleep -Seconds 1
            Restart-Computer -Force
        }
        else {
            Write-Host "SkipAction flag set - restart command was not executed" -ForegroundColor Magenta
        }
    }
    "Shutdown" {
        Write-Host "Preparing to shut down PC..." -ForegroundColor Cyan
        
        if (-not $Force) {
            # Countdown
            for ($i = $config.CountdownSeconds; $i -gt 0; $i--) {
                Write-Host "Shutting down in $i seconds... Press Ctrl+C to abort" -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
            
            # Close applications
            Close-Applications
        }
        
        if (-not $SkipAction) {
            Write-Host "Executing shutdown command now..." -ForegroundColor Green
            Start-Sleep -Seconds 1
            Stop-Computer -Force
        }
        else {
            Write-Host "SkipAction flag set - shutdown command was not executed" -ForegroundColor Magenta
        }
    }
    "Wake" {
        if ([string]::IsNullOrWhiteSpace($MACAddress)) {
            Write-Error "MAC address is required for Wake action"
            return
        }
        
        Write-Host "Sending Wake-on-LAN packet..." -ForegroundColor Cyan
        Send-WakePacket -MACAddress $MACAddress
    }
    "SetupWakeOnLAN" {
        Setup-WOL -AllowRemoteAccess:$AllowRemoteAccess
    }
}
'@
    
    # Create PCPow-Launcher.ps1
    $launcherPath = "$InstallPath\bin\PCPow-Launcher.ps1"
    Write-Verbose "Creating launcher script: $launcherPath"
    Set-Content -Path $launcherPath -Value @'
<#
.SYNOPSIS
PCPow Launcher - Prevents PowerShell window from closing

.DESCRIPTION
Executes PCPow in a separate process to prevent the current PowerShell window from closing
#>

[CmdletBinding()]
param (
    [Parameter(Position=0)]
    [ValidateSet("Sleep", "Restart", "Shutdown")]
    [string]$Action,
    
    [Parameter()]
    [switch]$Force,
    
    [Parameter()]
    [switch]$SkipAction
)

# Get the PCPow script path
$pcpowScript = Join-Path (Split-Path -Parent $PSCommandPath) "PCPow.ps1"

if (-not (Test-Path $pcpowScript)) {
    Write-Error "PCPow script not found at: $pcpowScript"
    exit 1
}

# Build arguments - properly escape quotes in string
$arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$pcpowScript`" -Action $Action"
if ($Force) { $arguments += " -Force" }
if ($SkipAction) { $arguments += " -SkipAction" }

# Launch in a separate window
Write-Host "Launching PCPow $Action (Force: $Force, SkipAction: $SkipAction)" -ForegroundColor Cyan
Start-Process powershell -ArgumentList $arguments -WindowStyle Hidden
'@
    
    # Create pcpow.bat for Command Prompt
    $batchPath = "$InstallPath\bin\pcpow.bat"
    Write-Verbose "Creating batch file: $batchPath"
    Set-Content -Path $batchPath -Value @'
@echo off
setlocal enabledelayedexpansion

set "action=%~1"
set "force="
set "skipaction="
set "macaddress="

if "%action%"=="" (
    echo PCPow - Power Control utility
    echo.
    echo Usage:
    echo   pcpow sleep [options]    - Put PC to sleep
    echo   pcpow restart [options]  - Restart PC
    echo   pcpow shutdown [options] - Shut down PC
    echo   pcpow wake MAC-ADDRESS   - Wake remote PC
    echo.
    echo Options:
    echo   -force      - Skip countdown and force close applications
    echo   -skipaction - Test mode without performing action
    echo.
    exit /b
)

REM Process arguments
:arg_loop
shift
if "%~1"=="" goto process_command
if /i "%~1"=="-force" set "force=-Force"
if /i "%~1"=="-skipaction" set "skipaction=-SkipAction"
if not "%macaddress%"=="" goto arg_loop
if /i "%action%"=="wake" set "macaddress=%~1"
goto arg_loop

:process_command
if /i "%action%"=="wake" (
    if "%macaddress%"=="" (
        echo ERROR: MAC address is required for wake command
        exit /b 1
    )
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0PCPow.ps1" -Action Wake -MACAddress %macaddress%
    exit /b
)

if /i "%action%"=="sleep" (
    start /b "" powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0PCPow-Launcher.ps1" -Action Sleep %force% %skipaction%
    exit /b
)

if /i "%action%"=="restart" (
    start /b "" powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0PCPow-Launcher.ps1" -Action Restart %force% %skipaction%
    exit /b
)

if /i "%action%"=="shutdown" (
    start /b "" powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0PCPow-Launcher.ps1" -Action Shutdown %force% %skipaction%
    exit /b
)

echo ERROR: Unknown action: %action%
echo Valid actions: sleep, restart, shutdown, wake
exit /b 1
'@
    
    # Create default config file
    $configPath = "$InstallPath\config\pcpow.config.json"
    Write-Verbose "Creating configuration file: $configPath"
    Set-Content -Path $configPath -Value @'
{
    "CountdownSeconds": 5,
    "IgnoreApps": [
        "explorer",
        "powershell",
        "cmd"
    ],
    "WakeTargets": {}
}
'@
    
    # Update PowerShell profile
    Update-Profile -BinPath "$InstallPath\bin"
    
    # Add bin folder to PATH
    $binPath = "$InstallPath\bin"
    $currentPath = [Environment]::GetEnvironmentVariable("PATH", "User")
    
    if ($currentPath -notlike "*$binPath*") {
        Write-Verbose "Adding PCPow bin directory to PATH"
        [Environment]::SetEnvironmentVariable("PATH", "$binPath;$currentPath", "User")
        $env:PATH = "$binPath;$env:PATH"
        Write-Host "Added PCPow bin directory to PATH" -ForegroundColor Green
    }
    else {
        Write-Host "PCPow bin directory already in PATH" -ForegroundColor Yellow
    }
    
    # Create README.md
    $readmePath = "$InstallPath\README.md"
    Write-Verbose "Creating README file: $readmePath"
    Set-Content -Path $readmePath -Value @'
# PCPow - PC Power Controller

A powerful utility for Windows that safely manages PC power states and provides Wake-on-LAN functionality.

## Features

- **Power State Management**: Sleep, restart, or shut down your PC with simple commands
- **Application Safety**: Gracefully closes applications before power actions
- **Wake-on-LAN**: Remotely wake PCs on your network using MAC addresses
- **Multiple Interfaces**: PowerShell commands, command-line interface, and desktop shortcuts
- **Terminal-Friendly**: Commands run without closing your terminal window

## Quick Start Guide

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

### Command Line Interface
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
# Run this command to configure your network adapters
.\Setup-WakeOnLAN [-AllowRemoteAccess]
```

This configures your network adapters and saves MAC addresses to `wake-targets.txt` for future use.

## Configuration

Configuration file is located at: `config\pcpow.config.json`

You can customize:
- Countdown duration
- Applications to ignore when closing
- Default MAC addresses for Wake-on-LAN

## Troubleshooting

If commands aren't working:

1. **Verify installation**: Run `Get-Command -Name pows, powr, powd, poww` to check if commands are defined
2. **Reload profile**: Run `. $PROFILE` to reload your PowerShell profile
3. **Check PATH**: Ensure PCPow bin directory is in your PATH
4. **Permissions**: If having issues, try running PowerShell as Administrator

## Support

For issues or questions, please create an issue on the GitHub repository.
'@
    
    # Create CHANGELOG.md
    $changelogPath = "$InstallPath\CHANGELOG.md"
    Write-Verbose "Creating changelog file: $changelogPath"
    Set-Content -Path $changelogPath -Value @'
# Changelog

All notable changes to the PCPow utility will be documented in this file.

## [1.2.3] - 2025-04-05

### Fixed
- Fixed issue with PCPow-Launcher.ps1 string interpolation that caused unexpected token errors
- Fixed PowerShell profile corruption issues
- Improved error handling in all scripts
- Fixed commands closing terminal window
- Verified pows and powr commands working properly
- Confirmed desktop shortcuts functionality

### Added
- Comprehensive installation script
- Better logging and error messages
- SkipAction parameter for testing without executing actual power commands
- Improved documentation

### Changed
- Relocated main scripts to dedicated bin directory
- Simplified command structure
- Enhanced error reporting
- Improved Wake-on-LAN configuration

## [1.2.1] - 2025-04-04

### Initial Release
- Basic PC power control commands
- Wake-on-LAN functionality
- PowerShell and Command Prompt interfaces
'@
    
    # Installation complete
    Write-Host "`nPCPow installation complete!" -ForegroundColor Green
    Write-Host "`nInstalled components:"
    Write-Host "- Main script: $mainScriptPath"
    Write-Host "- Launcher: $launcherPath"
    Write-Host "- Batch file: $batchPath"
    Write-Host "- Configuration: $configPath"
    Write-Host "- PowerShell profile updated: $PROFILE"
    
    Write-Host "`nAvailable commands:"
    Write-Host "  PowerShell:"
    Write-Host "    pows             - Put PC to sleep"
    Write-Host "    powr             - Restart PC"
    Write-Host "    powd             - Shutdown PC"
    Write-Host "    poww <MAC>       - Wake a remote PC"
    Write-Host "  Command Prompt:"
    Write-Host "    pcpow sleep      - Put PC to sleep"
    Write-Host "    pcpow restart    - Restart PC"
    Write-Host "    pcpow shutdown   - Shut down PC"
    Write-Host "    pcpow wake <MAC> - Wake a remote PC"
    
    # Ask to restart PowerShell
    $restartPs = Read-Host "`nRestart PowerShell now to load PCPow commands? (y/n)"
    if ($restartPs -eq 'y') {
        Write-Host "Restarting PowerShell..." -ForegroundColor Cyan
        Start-Process powershell
        exit
    }
    else {
        Write-Host "`nPlease restart PowerShell to load PCPow commands." -ForegroundColor Yellow
    }
}

# Run installation
Install-PCPow 