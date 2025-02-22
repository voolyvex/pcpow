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

function Show-ConfirmationPrompt {
    param (
        [string]$ActionType
    )

    $title = "Confirm $ActionType"
    $message = "WARNING: This will close all running programs and $ActionType.`nSave any important work before continuing.`nProceed?"
    
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Confirm $ActionType"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Cancel operation"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    
    $result = $host.ui.PromptForChoice($title, $message, $options, 1)
    return $result -eq 0
}

function Get-UserApps {
    $excludedProcesses = @(
        "explorer", "svchost", "csrss", "smss", "wininit", "winlogon",
        "spoolsv", "lsass", "services", "system", "registry", "idle",
        "dwm", "RuntimeBroker", "ShellExperienceHost", "SearchHost",
        "StartMenuExperienceHost", "Taskmgr", "sihost", "fontdrvhost",
        "powershell", "conhost", "WmiPrvSE", "dllhost", "ctfmon",
        "SecurityHealthService", "SearchIndexer", "Memory Compression"
    )
    
    $processes = Get-Process | Where-Object {
        $_.SessionId -eq (Get-Process -PID $PID).SessionId -and
        $_.Name -notin $excludedProcesses -and
        $_.MainWindowHandle -ne 0
    }
    
    return $processes
}

function Close-Apps {
    param (
        [Parameter(Mandatory=$true)]
        [System.Diagnostics.Process[]]$Processes,
        [switch]$Force,
        [string]$ActionType
    )
    
    $failedProcesses = @()
    foreach ($proc in $Processes) {
        try {
            Write-Host "Attempting to close $($proc.Name)..." -ForegroundColor Yellow
            if ($proc.CloseMainWindow()) {
                # Wait up to 5 seconds for the process to close gracefully
                if (!$proc.WaitForExit(5000)) {
                    if ($Force) {
                        Write-Host "Force closing $($proc.Name)..." -ForegroundColor Red
                        $proc | Stop-Process -Force
                    } else {
                        $failedProcesses += $proc.Name
                    }
                } else {
                    Write-Host "$($proc.Name) closed successfully" -ForegroundColor Green
                }
            } else {
                if ($Force) {
                    Write-Host "Force closing $($proc.Name)..." -ForegroundColor Red
                    $proc | Stop-Process -Force
                } else {
                    $failedProcesses += $proc.Name
                }
            }
        } catch {
            Write-Warning "Failed to close $($proc.Name): $_"
            $failedProcesses += $proc.Name
        }
    }
    
    if ($failedProcesses.Count -gt 0) {
        Write-Warning "The following applications could not be closed:`n$($failedProcesses -join "`n")"
        if (-not $Force) {
            Write-Host "Try running the script with -Force to forcefully close applications" -ForegroundColor Yellow
        }
    }
}

try {
    if (-not $Force) {
        if (-not (Show-ConfirmationPrompt -ActionType "restart")) {
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
        if (-not (Close-Apps -Processes $userApps -Force:$Force -ActionType "restart")) {
            throw "Failed to close all applications"
        }
    }

    Write-PCPowLog "Waiting for processes to finish closing..." -Level Info
    Start-Sleep -Seconds 2

    Write-PCPowLog "Restarting PC..." -Level Action
    Invoke-PowerAction -Action Restart -Force:$Force
}
catch {
    Write-PCPowLog "Error: $_" -Level Error
    exit 1
} 