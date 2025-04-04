# Setup script to install shortcuts and aliases
$ErrorActionPreference = 'Stop'

# Get the directory where the scripts are located
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = [System.IO.Path]::GetFullPath($scriptPath)

Write-Host "PCPow Setup" -ForegroundColor Cyan
Write-Host "-----------" -ForegroundColor Cyan
Write-Host ""

# Create shortcuts directory in the Windows directory if it doesn't exist
$shortcutsDir = Join-Path $env:USERPROFILE "AppData\Local\Microsoft\WindowsApps"
if (-not (Test-Path $shortcutsDir)) {
    Write-Host "Creating shortcuts directory: $shortcutsDir" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $shortcutsDir -Force | Out-Null
}

# Copy config file
Write-Host "Copying configuration file..." -ForegroundColor Green
Copy-Item "$scriptPath\pcpow.config.json" "$shortcutsDir\" -Force
if (-not (Test-Path "$shortcutsDir\pcpow.config.json")) {
    throw "Failed to copy config file to $shortcutsDir\pcpow.config.json"
}

# Copy script files
Write-Host "Copying script files..." -ForegroundColor Green
$filesToCopy = @(
    "pcpow.bat",
    "PCPow.ps1",
    "Setup-WakeOnLAN.ps1",
    "update-profile.ps1"
)

foreach ($file in $filesToCopy) {
    Write-Host "  Copying $file" -ForegroundColor White
    Copy-Item "$scriptPath\$file" $shortcutsDir -Force
    if (-not (Test-Path "$shortcutsDir\$file")) {
        throw "Failed to copy $file to $shortcutsDir"
    }
}

# Create PowerShell profile directory if it doesn't exist
$profileDir = Split-Path -Parent $PROFILE
if (-not (Test-Path $profileDir)) {
    Write-Host "Creating PowerShell profile directory: $profileDir" -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $profileDir -Force | Out-Null
}

# Create PowerShell profile if it doesn't exist
if (-not (Test-Path $PROFILE)) {
    Write-Host "Creating PowerShell profile: $PROFILE" -ForegroundColor Yellow
    New-Item -ItemType File -Path $PROFILE -Force | Out-Null
}

# Use the update-profile.ps1 script to update the profile
Write-Host "Updating PowerShell profile using update-profile.ps1..." -ForegroundColor Green
$updateProfilePath = Join-Path $shortcutsDir "update-profile.ps1"
if (Test-Path $updateProfilePath) {
    Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$updateProfilePath`"" -Wait -NoNewWindow
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Profile updater returned non-zero exit code: $LASTEXITCODE"
        Write-Warning "You may need to run update-profile.ps1 manually."
    }
} else {
    Write-Warning "update-profile.ps1 not found at $updateProfilePath. Profile may need manual update."
}

# Test the commands
Write-Host "`nTesting command availability..." -ForegroundColor Green
$testScript = @"
try {
    Write-Host "Verifying PCPow installation..."
    if (Test-Path "$shortcutsDir\PCPow.ps1") {
        Write-Host "  PCPow.ps1: Found" -ForegroundColor Green
    } else {
        Write-Host "  PCPow.ps1: Not found" -ForegroundColor Red
    }
    
    if (Test-Path "$shortcutsDir\pcpow.bat") {
        Write-Host "  pcpow.bat: Found" -ForegroundColor Green
    } else {
        Write-Host "  pcpow.bat: Not found" -ForegroundColor Red
    }
    
    Write-Host "`nAvailable commands:"
    Write-Host "  pcpow sleep/restart/shutdown/wake"
    Write-Host "  pows, powr, powd, poww"
    Write-Host "  Start-PCSleep, Restart-PC, Stop-PC, Wake-PC"
} catch {
    Write-Warning "Command verification failed: $($_)"
}
"@

powershell -NoProfile -Command $testScript

# Check for old module directory
$moduleVersions = @("1.0.0", "1.1.0")
foreach ($ver in $moduleVersions) {
    $oldModulePath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "WindowsPowerShell\Modules\pcpow-common\$ver"
    if (Test-Path $oldModulePath) {
        Write-Host "  Removing old module directory: $oldModulePath" -ForegroundColor White
        Remove-Item $oldModulePath -Force -Recurse
    }
}

Write-Host "`nSetup completed successfully!" -ForegroundColor Green
Write-Host @"

PCPow v1.2.1 has been installed:
1. Scripts installed to: $shortcutsDir
2. Profile updated at: $PROFILE

You can now use the following commands from anywhere:
1. From Run menu (Win+R) or Command Prompt:
   pcpow sleep     # Sleep PC
   pcpow restart   # Restart PC
   pcpow shutdown  # Shutdown PC
   pcpow wake MAC  # Wake a remote PC

2. From PowerShell:
   Quick aliases:
   pows            # Sleep
   powr            # Restart
   powd            # Shutdown
   poww MAC        # Wake PC
   
   Full commands:
   Start-PCSleep
   Restart-PC
   Stop-PC
   Wake-PC MAC

Add -Force to any command to skip confirmation and force close apps.
Example: pcpow sleep -Force or pows -Force

To setup Wake-on-LAN for this PC, run:
   Setup-WakeOnLAN.ps1 -AllowRemoteAccess

Please close and reopen your PowerShell window for the changes to take effect.
"@

# Prompt to restart PowerShell
$restart = Read-Host "Would you like to restart PowerShell now? (Y/N)"
if ($restart -eq 'Y' -or $restart -eq 'y') {
    Start-Process powershell
    exit
} 