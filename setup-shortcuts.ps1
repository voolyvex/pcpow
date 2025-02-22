# Setup script to install shortcuts and aliases
$ErrorActionPreference = 'Stop'

# Get the directory where the scripts are located
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Create shortcuts directory in the Windows directory if it doesn't exist
$shortcutsDir = "$env:USERPROFILE\AppData\Local\Microsoft\WindowsApps"
if (-not (Test-Path $shortcutsDir)) {
    New-Item -ItemType Directory -Path $shortcutsDir -Force
}

# Copy the batch file to the shortcuts directory
Copy-Item "$scriptPath\pcpow.bat" $shortcutsDir -Force

# Create PowerShell profile if it doesn't exist
if (-not (Test-Path $PROFILE)) {
    New-Item -ItemType File -Path $PROFILE -Force
}

# Add aliases to PowerShell profile
$aliasContent = @"

# PCPow - Power Management Aliases
Set-Alias -Name pow -Value "$shortcutsDir\pcpow.bat"
function Sleep-PC { & "$shortcutsDir\pcpow.bat" sleep `$args }
function Restart-PCApps { & "$shortcutsDir\pcpow.bat" restart `$args }
function Stop-PCApps { & "$shortcutsDir\pcpow.bat" shutdown `$args }

# Create shorter aliases
Set-Alias -Name pows -Value Sleep-PC      # pow sleep
Set-Alias -Name powr -Value Restart-PCApps # pow restart
Set-Alias -Name powd -Value Stop-PCApps    # pow shutdown
"@

Add-Content -Path $PROFILE -Value $aliasContent

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

Note: You'll need to restart your PowerShell session for the aliases to take effect.
"@ 