# Requires -Version 5.1
<#
.SYNOPSIS
Safely closes applications and puts the computer to sleep

.DESCRIPTION
This script gracefully closes all user applications before initiating sleep mode.
It uses configurable timeouts and process exclusion lists from pcpow.config.json.

.PARAMETER Force
If specified, forces applications to close without waiting for them to exit gracefully.

.EXAMPLE
.\Close-AndSleep.ps1
Prompts for confirmation before closing applications and sleeping

.EXAMPLE
.\Close-AndSleep.ps1 -Force
Forces applications to close and sleeps without confirmation

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
    if ($Force -or (Show-ConfirmationPrompt -ActionType "Sleep")) {
        Write-PCPowLog "Initiating sleep sequence..." -Level Action
        if (Close-Applications -TimeoutMS $script:Config.TimeoutMS) {
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.Application]::SetSuspendState("Suspend", $false, $false)
        }
    }
}
catch {
    Write-PCPowLog "Error: $_" -Level Error
    exit 1
} 