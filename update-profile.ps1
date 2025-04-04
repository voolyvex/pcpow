# update-profile.ps1
# Run this script to update your PowerShell profile with the correct PCPow configuration.
# Execute with: powershell -ExecutionPolicy Bypass -File .\update-profile.ps1

$ErrorActionPreference = 'Stop'

# --- Configuration ---
$ProfilePath = "$env:USERPROFILE\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1"
$ShortcutsDir = Join-Path $env:USERPROFILE "AppData\Local\Microsoft\WindowsApps" # PCPow install location
# --- End Configuration ---

Write-Host "Updating PowerShell Profile: $ProfilePath" -ForegroundColor Cyan
Write-Host "--------------------------------------" -ForegroundColor Cyan

# Verify PCPow script exists in the expected location
$PCPowScriptPath = Join-Path $ShortcutsDir "PCPow.ps1"
if (-not (Test-Path $PCPowScriptPath)) {
    Write-Error "PCPow.ps1 not found at '$PCPowScriptPath'. Please ensure PCPow is installed correctly using setup-shortcuts.ps1."
    exit 1
}
$PCPowBatPath = Join-Path $ShortcutsDir "pcpow.bat"
if (-not (Test-Path $PCPowBatPath)) {
    Write-Error "pcpow.bat not found at '$PCPowBatPath'. Please ensure PCPow is installed correctly using setup-shortcuts.ps1."
    exit 1
}


# --- Correct PCPow Profile Block ---
# Note: Using backticks ` to escape $ within the here-string
$CorrectPCPowBlock = @"
# PCPow Start
# PCPow Configuration v1.2.1 (Managed by update-profile.ps1)
`$ErrorActionPreference = 'Stop'

# PCPow - Power Management Functions
function global:Start-PCSleep {
    param(
        [switch]`$Force,
        [switch]`$SkipSleep
    )

    `$scriptPath = "$ShortcutsDir\PCPow.ps1"

    if (Test-Path `$scriptPath) {
        # Use parameter hashtable for proper splatting
        `$params = @{
            Action = "Sleep"
        }

        if (`$Force) { `$params.Force = `$true }
        if (`$SkipSleep) { `$params.SkipAction = `$true }

        & `$scriptPath @params
    } else {
        Write-Warning "PCPow script not found at `$scriptPath"
    }
}

function global:Restart-PC {
    param(
        [switch]`$Force,
        [switch]`$SkipRestart
    )

    `$scriptPath = "$ShortcutsDir\PCPow.ps1"

    if (Test-Path `$scriptPath) {
        # Use parameter hashtable for proper splatting
        `$params = @{
            Action = "Restart"
        }

        if (`$Force) { `$params.Force = `$true }
        if (`$SkipRestart) { `$params.SkipAction = `$true }

        & `$scriptPath @params
    } else {
        Write-Warning "PCPow script not found at `$scriptPath"
    }
}

function global:Stop-PC {
    param(
        [switch]`$Force,
        [switch]`$SkipShutdown
    )

    `$scriptPath = "$ShortcutsDir\PCPow.ps1"

    if (Test-Path `$scriptPath) {
        # Use parameter hashtable for proper splatting
        `$params = @{
            Action = "Shutdown"
        }

        if (`$Force) { `$params.Force = `$true }
        if (`$SkipShutdown) { `$params.SkipAction = `$true }

        & `$scriptPath @params
    } else {
        Write-Warning "PCPow script not found at `$scriptPath"
    }
}

function global:Wake-PC {
    param(
        [Parameter(Mandatory=`$true)]
        [string]`$MACAddress
    )

    `$scriptPath = "$ShortcutsDir\pcpow.bat"

    if (Test-Path `$scriptPath) {
        & `$scriptPath wake `$MACAddress
    } else {
        Write-Warning "PCPow batch file not found at `$scriptPath"
    }
}

# Create aliases
Set-Alias -Name pows -Value Start-PCSleep -Scope Global
Set-Alias -Name powr -Value Restart-PC -Scope Global
Set-Alias -Name powd -Value Stop-PC -Scope Global
Set-Alias -Name poww -Value Wake-PC -Scope Global

Write-Host "PCPow v1.2.1 commands loaded successfully" -ForegroundColor Green
# PCPow End
"@
# --- End Correct PCPow Profile Block ---

try {
    # Create profile directory if it doesn't exist
    $ProfileDir = Split-Path -Parent $ProfilePath
    if (-not (Test-Path $ProfileDir -PathType Container)) {
        Write-Host "Creating PowerShell profile directory: $ProfileDir" -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $ProfileDir -Force | Out-Null
    }

    # Backup existing profile, if it exists
    if (Test-Path $ProfilePath -PathType Leaf) {
        $backupPath = "$ProfilePath.bak.$(Get-Date -Format 'yyyyMMddHHmmss')"
        Write-Host "Backing up existing profile to: $backupPath" -ForegroundColor Yellow
        Copy-Item -Path $ProfilePath -Destination $backupPath -Force
    }

    # Start with a fresh profile - completely replace rather than trying to clean up
    Write-Host "Creating clean profile at: $ProfilePath" -ForegroundColor Green
    
    # Basic profile with PCPow configuration
    $newProfileContent = @"
# PowerShell Profile Configuration
# Created/Updated by PCPow update-profile.ps1 script
# Version 1.2.1
# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

$CorrectPCPowBlock

# Add your additional PowerShell customizations below this line
"@

    # Write the new content
    Write-Host "Writing updated configuration to profile..." -ForegroundColor Green
    Set-Content -Path $ProfilePath -Value $newProfileContent -Encoding UTF8 -Force

    Write-Host "`nProfile completely rebuilt!" -ForegroundColor Green
    Write-Host "Please close and reopen any PowerShell windows for the changes to take effect." -ForegroundColor Yellow

} catch {
    Write-Error "Failed to update profile: $_"
    if (Test-Path $backupPath) {
        Write-Warning "Your original profile is backed up at: $backupPath"
    }
    exit 1
} 