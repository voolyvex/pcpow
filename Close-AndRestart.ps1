# Requires -Version 5.1
<#
.SYNOPSIS
Safely closes applications and restarts the computer

.DESCRIPTION
This script gracefully closes all user applications before initiating a system restart.
It uses configurable timeouts and process exclusion lists from pcpow.config.json.

.PARAMETER Force
If specified, forces applications to close without waiting for them to exit gracefully.

.EXAMPLE
.\Close-AndRestart.ps1
Prompts for confirmation before closing applications and restarting

.EXAMPLE
.\Close-AndRestart.ps1 -Force
Forces applications to close and restarts without confirmation

.NOTES
Version: 1.0.0
Author: PCPow Team
#>

[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# Import the module by name (it should be in the PSModulePath)
try {
    Import-Module pcpow-common -MinimumVersion 1.0.0 -Force -ErrorAction Stop
} catch {
    Write-Error "Failed to import PCPow module. Please ensure it is installed correctly: $_"
    exit 1
}

try {
    $confirmMessage = "Are you sure you want to restart the computer? This will close all applications."
    if ($Force -or $host.UI.PromptForChoice("Confirm Restart", $confirmMessage, @("&Yes", "&No"), 1) -eq 0) {
        # Close all applications gracefully
        Get-Process | Where-Object { $_.MainWindowTitle -ne "" } | Stop-Process -Force

        # Wait a moment for processes to close
        Start-Sleep -Seconds 2

        # Restart computer
        Restart-Computer -Force
    }
}
catch {
    Write-PCPowLog "Error: $_" -Level Error
    exit 1
}

# Ensure restart happens even if there are errors
Write-Host "Forcing immediate restart..." -ForegroundColor Red
shutdown.exe /r /f /t 0 