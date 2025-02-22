# Requires -Version 5.1
<#
.SYNOPSIS
Shared functionality for PCPow power management scripts

.DESCRIPTION
This module provides common functions and configuration management for PCPow scripts.
It handles process management, user interaction, and power actions with configurable settings.

.NOTES
Version: 1.0.0
Author: PCPow Team
#>

# Initialize configuration
$script:Config = $null

function Initialize-PCPowConfig {
    [CmdletBinding()]
    param(
        [string]$ConfigPath = "$PSScriptRoot\pcpow.config.json"
    )
    
    try {
        if (Test-Path $ConfigPath) {
            $script:Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        } else {
            Write-Warning "Configuration file not found at $ConfigPath. Using default settings."
            $script:Config = @{
                version = "1.0.0"
                timeoutMS = 5000
                colors = @{
                    warning = "Yellow"
                    success = "Green"
                    error = "Red"
                    info = "Cyan"
                    action = "Magenta"
                }
                excludedProcesses = @(
                    "explorer", "svchost", "csrss", "smss", "wininit", "winlogon",
                    "spoolsv", "lsass", "services", "system", "registry", "idle"
                )
            }
        }
    }
    catch {
        Write-Error "Failed to initialize configuration: $_"
        throw
    }
}

function Write-PCPowLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [ValidateSet('Info','Success','Warning','Error','Action')]
        [string]$Level = 'Info'
    )
    
    $color = $script:Config.colors."$($Level.ToLower())"
    Write-Host $Message -ForegroundColor $color
}

function Get-UserApps {
    [CmdletBinding()]
    param()
    
    try {
        Get-Process | Where-Object {
            $_.SessionId -eq (Get-Process -PID $PID).SessionId -and
            $_.Name -notin $script:Config.excludedProcesses -and
            $_.MainWindowHandle -ne 0
        }
    }
    catch {
        Write-PCPowLog "Failed to get user applications: $_" -Level Error
        throw
    }
}

function Close-Apps {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.Diagnostics.Process[]]$Processes,
        [switch]$Force,
        [string]$ActionType
    )

    $failedProcesses = @()
    foreach ($proc in $Processes) {
        try {
            Write-Host "Attempting to close $($proc.Name)..." -ForegroundColor $script:Config.Colors.Info
            if ($proc.CloseMainWindow()) {
                if (!$proc.WaitForExit($script:Config.TimeoutMS)) {
                    if ($Force) {
                        Write-Host "Force closing $($proc.Name)..." -ForegroundColor $script:Config.Colors.Error
                        $proc | Stop-Process -Force
                    } else {
                        $failedProcesses += $proc.Name
                    }
                } else {
                    Write-Host "$($proc.Name) closed successfully" -ForegroundColor $script:Config.Colors.Success
                }
            } else {
                if ($Force) {
                    Write-Host "Force closing $($proc.Name)..." -ForegroundColor $script:Config.Colors.Error
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
            Write-Host "Try running with -Force to override" -ForegroundColor $script:Config.Colors.Warning
        }
        return $false
    }
    return $true
}

function Show-ConfirmationPrompt {
    [CmdletBinding()]
    param(
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

function Invoke-PowerAction {
    [CmdletBinding()]
    param(
        [ValidateSet('Restart','Shutdown','Sleep')]
        [string]$Action,
        [switch]$Force
    )

    try {
        Write-Host "Initiating $Action sequence..." -ForegroundColor $script:Config.Colors.Action
        
        $userApps = Get-UserApps
        if ($userApps.Count -gt 0) {
            Write-Host "Closing $($userApps.Count) applications..." -ForegroundColor $script:Config.Colors.Info
            if (-not (Close-Apps -Processes $userApps -Force:$Force -ActionType $Action)) {
                throw "Application closure failed"
            }
        }

        Write-Host "Performing $Action..." -ForegroundColor $script:Config.Colors.Action
        switch ($Action) {
            'Restart' { Restart-Computer -Force:$Force }
            'Shutdown' { Stop-Computer -Force:$Force }
            'Sleep' { Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Application]::SetSuspendState("Suspend", $false, $false) }
        }
    }
    catch {
        Write-Host "Error during $Action operation: $_" -ForegroundColor $script:Config.Colors.Error
        exit 1
    }
}

# Initialize configuration when module is imported
Initialize-PCPowConfig

Export-ModuleMember -Function @(
    'Initialize-PCPowConfig',
    'Get-UserApps',
    'Close-Apps',
    'Show-ConfirmationPrompt',
    'Invoke-PowerAction',
    'Write-PCPowLog'
) 