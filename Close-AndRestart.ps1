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
    # Call Invoke-PowerAction which handles all the logic including confirmation
    Invoke-PowerAction -Action 'Restart' -Force:$Force
}
catch {
    Write-PCPowLog "Error during restart operation: $_" -Level Error
    exit 1
} 