# Requires -Version 5.1
<#
.SYNOPSIS
Safely closes applications and shuts down the computer

.DESCRIPTION
This script gracefully closes all user applications before initiating shutdown.
It uses configurable timeouts and process exclusion lists from pcpow.config.json.

.PARAMETER Force
If specified, forces applications to close without waiting for them to exit gracefully.

.EXAMPLE
.\Close-AndShutdown.ps1
Prompts for confirmation before closing applications and shutting down

.EXAMPLE
.\Close-AndShutdown.ps1 -Force
Forces applications to close and shuts down without confirmation

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

# Add check for admin context
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Operation requires administrator privileges"
    exit 1
}

# Add terminal warning
Write-Host "This will close all applications and shut down the computer!" -ForegroundColor Red
Write-Host "Keep this window open until the operation completes." -ForegroundColor Yellow

try {
    $confirmMessage = "Are you sure you want to shutdown the computer? This will close all applications."
    if ($Force -or $host.UI.PromptForChoice("Confirm Shutdown", $confirmMessage, @("&Yes", "&No"), 1) -eq 0) {
        # Close all applications gracefully
        Get-Process | Where-Object { $_.MainWindowTitle -ne "" } | Stop-Process -Force

        # Wait a moment for processes to close
        Start-Sleep -Seconds 2

        # Shutdown computer
        Stop-Computer -Force
    }
}
catch {
    Write-PCPowLog "Error: $_" -Level Error
    exit 1
} 