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
    if (-not $Force) {
        if (-not (Show-ConfirmationPrompt -ActionType "sleep")) {
            Write-PCPowLog "Operation cancelled by user." -Level Warning
            exit 0
        }
    }

    Write-PCPowLog "Identifying running applications..." -Level Info
    $userApps = Get-UserApps

    if ($userApps.Count -eq 0) {
        Write-PCPowLog "No user applications found to close." -Level Success
    } else {
        Write-PCPowLog "Found $($userApps.Count) applications to close." -Level Info
        if (-not (Close-Apps -Processes $userApps -Force:$Force -ActionType "sleep")) {
            throw "Failed to close all applications"
        }
    }

    Write-PCPowLog "Waiting for processes to finish closing..." -Level Info
    Start-Sleep -Seconds 2

    Write-PCPowLog "Putting PC to sleep..." -Level Action
    Invoke-PowerAction -Action Sleep -Force:$Force
}
catch {
    Write-PCPowLog "Error: $_" -Level Error
    exit 1
} 