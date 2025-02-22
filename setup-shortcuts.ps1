# Setup script to install shortcuts and aliases
$ErrorActionPreference = 'Stop'

# Get the directory where the scripts are located
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = [System.IO.Path]::GetFullPath($scriptPath)

# Create shortcuts directory in the Windows directory if it doesn't exist
$shortcutsDir = "$env:USERPROFILE\AppData\Local\Microsoft\WindowsApps"
if (-not (Test-Path $shortcutsDir)) {
    New-Item -ItemType Directory -Path $shortcutsDir -Force
}

# Copy all necessary files
Write-Host "Copying files to $shortcutsDir..."
Copy-Item "$scriptPath\pcpow.bat" $shortcutsDir -Force
Copy-Item "$scriptPath\Close-And*.ps1" $shortcutsDir -Force

# Create PowerShell profile directory if it doesn't exist
$profileDir = Split-Path -Parent $PROFILE
if (-not (Test-Path $profileDir)) {
    New-Item -ItemType Directory -Path $profileDir -Force
}

# Create PowerShell profile if it doesn't exist
if (-not (Test-Path $PROFILE)) {
    New-Item -ItemType File -Path $PROFILE -Force
}

# Add module path to profile
$moduleConfig = @"
# PCPow Configuration
`$env:PSModulePath += ";$shortcutsDir"
Import-Module pcpow-common -Force
"@

# Update profile content
$profileContent = @"
$moduleConfig

# PCPow - Power Management Aliases
function Sleep-PC { & "$shortcutsDir\Close-AndSleep.ps1" @args }
function Restart-PCApps { & "$shortcutsDir\Close-AndRestart.ps1" @args }
function Stop-PCApps { & "$shortcutsDir\Close-AndShutdown.ps1" @args }

# Short aliases
Set-Alias -Name pows -Value Sleep-PC
Set-Alias -Name powr -Value Restart-PCApps
Set-Alias -Name powd -Value Stop-PCApps
"@

# Remove old aliases if they exist
$currentContent = Get-Content $PROFILE -ErrorAction SilentlyContinue
if ($currentContent) {
    $currentContent = $currentContent | Where-Object { $_ -notmatch "PCPow - Power Management" }
    Set-Content $PROFILE $currentContent
}

Add-Content -Path $PROFILE -Value $profileContent

# Update pcpow.bat to use full paths
$batchContent = @"
@echo off
if "%~1"=="" goto :help
if /i "%~1"=="sleep" goto :sleep
if /i "%~1"=="restart" goto :restart
if /i "%~1"=="shutdown" goto :shutdown
if /i "%~1"=="-h" goto :help
if /i "%~1"=="--help" goto :help
goto :help

:sleep
powershell -ExecutionPolicy Bypass -File "$shortcutsDir\Close-AndSleep.ps1" %2
goto :eof

:restart
powershell -ExecutionPolicy Bypass -File "$shortcutsDir\Close-AndRestart.ps1" %2
goto :eof

:shutdown
powershell -ExecutionPolicy Bypass -File "$shortcutsDir\Close-AndShutdown.ps1" %2
goto :eof

:help
echo PCPow - Windows Power Management
echo -------------------------------
echo Usage: pcpow [command] [-Force]
echo.
echo Commands:
echo   sleep     - Close all apps and put PC to sleep
echo   restart   - Close all apps and restart PC
echo   shutdown  - Close all apps and shutdown PC
echo.
echo Options:
echo   -Force    - Skip confirmation and force close apps
echo.
echo Examples:
echo   pcpow sleep
echo   pcpow restart -Force
echo   pcpow shutdown
"@

Set-Content -Path "$shortcutsDir\pcpow.bat" -Value $batchContent

Write-Host "Setup completed successfully!" -ForegroundColor Green
Write-Host @"

You can now use the following commands from anywhere:
1. From Run menu (Win+R) or Command Prompt:
   - pcpow sleep
   - pcpow restart
   - pcpow shutdown

2. From PowerShell:
   Short commands:
   - pow sleep   (or pows)
   - pow restart (or powr)
   - pow shutdown (or powd)
   
   Full commands:
   - Sleep-PC
   - Restart-PCApps
   - Stop-PCApps

Add -Force to any command to skip confirmation and force close apps.
Example: pow sleep -Force

Please close and reopen your PowerShell window for the changes to take effect.
"@

# Prompt to restart PowerShell
$restart = Read-Host "Would you like to restart PowerShell now? (Y/N)"
if ($restart -eq 'Y' -or $restart -eq 'y') {
    Start-Process powershell
    exit
} 