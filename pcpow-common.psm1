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
                    "svchost", "csrss", "smss", "wininit", "winlogon",
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

function Close-Applications {
    param([int]$TimeoutMS = 5000)
    
    # Track processes that need force closing
    [System.Collections.Generic.List[string]]$remainingProcesses = @()
    $processes = Get-Process | Where-Object {
        $_.MainWindowHandle -ne 0 -and 
        $_.ProcessName -notin $script:Config.excludedProcesses
    } | Sort-Object -Property WorkingSet64 -Descending
    
    # Get all PowerShell windows except current
    $powershellProcesses = Get-Process powershell, pwsh | Where-Object { $_.Id -ne $PID }
    $processes = @($processes) + @($powershellProcesses)

    # Include explorer.exe for restart/shutdown
    $explorer = Get-Process explorer -ErrorAction SilentlyContinue
    if ($explorer) { $processes += $explorer }
    
    if ($processes.Count -eq 0) {
        Write-PCPowLog "No applications to close." -Level Info
        return $true # No apps to close, consider it successful
    }

    Write-PCPowLog "Attempting to close $($processes.Count) applications..." -Level Info
    
    # Add timeout tracking
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    foreach ($process in $processes) {
        if ($stopwatch.ElapsedMilliseconds -ge $TimeoutMS) {
            Write-PCPowLog "Timeout reached, proceeding to force close remaining applications." -Level Warning
            break
        }
        try {
            Write-PCPowLog "Attempting to close $($process.Name) (PID: $($process.Id))..." -Level Info
            # Request process to close gracefully
            if ($process.CloseMainWindow()) {
                Write-PCPowLog "$($process.Name) (PID: $($process.Id)) main window closed, waiting for exit..." -Level Info
                $process.WaitForExit(($TimeoutMS - $stopwatch.ElapsedMilliseconds))
                if ($process.HasExited) {
                    Write-PCPowLog "$($process.Name) (PID: $($process.Id)) closed gracefully." -Level Success
                    continue # Proceed to next process
                } else {
                    Write-PCPowLog "$($process.Name) (PID: $($process.Id)) did not exit in time after MainWindowClose." -Level Warning
                    $remainingProcesses.Add($process.Name)
                }
            } else {
                Write-PCPowLog "Failed to send MainWindowClose to $($process.Name) (PID: $($process.Id))." -Level Warning
                $remainingProcesses.Add($process.Name)
            }
        } catch {
            Write-PCPowLog "Error closing $($process.Name): $_" -Level Warning
            $remainingProcesses.Add($process.Name)
        }
    }

    # Force close remaining processes
    if ($remainingProcesses.Count -gt 0) {
        Write-PCPowLog "Force closing $($remainingProcesses.Count) remaining applications..." -Level Warning
        foreach ($procName in $remainingProcesses) {
            try {
                # Use taskkill as a more forceful method
                Write-PCPowLog "Using taskkill /F to force close $($procName)..." -Level Warning
                taskkill /F /IM "$procName.exe" /T
                Write-PCPowLog "taskkill /F for $($procName) completed." -Level Success
            }
            catch {
                Write-PCPowLog "Failed to force close $($procName) using taskkill: $_" -Level Error
            }
        }
    }

    if ($remainingProcesses.Count -eq 0) {
        Write-PCPowLog "All applications closed successfully." -Level Success
    } else {
        Write-PCPowLog "Warning: Some applications may not have closed cleanly." -Level Warning
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
            if (-not (Close-Applications -TimeoutMS $script:Config.TimeoutMS)) {
                if ($Force) {
                    Write-PCPowLog "Forcefully terminating remaining applications..." -Level Warning
                    Get-Process | Where-Object {
                        $_.Id -in $remainingProcesses.Id
                    } | Stop-Process -Force
                }
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

function Start-PowerOperation {
    param(
        [ValidateSet('Sleep','Restart','Shutdown')]
        [string]$Operation,
        [switch]$Force
    )
    
    # Register cleanup handler
    Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
        Write-Warning "Operation interrupted! System state may be unstable."
        [System.Environment]::Exit(1)
    }
    
    try {
        # Existing logic...
    }
    catch {
        Write-Error "Operation failed: $_"
        exit 2
    }
}

# Initialize configuration when module is imported
Initialize-PCPowConfig

Export-ModuleMember -Function @(
    'Initialize-PCPowConfig',
    'Get-UserApps',
    'Sleep-PC',
    'Restart-PCApps',
    'Stop-PCApps',
    'Close-Applications',
    'Show-ConfirmationPrompt',
    'Invoke-PowerAction',
    'Write-PCPowLog'
) -Alias @('pows', 'powr', 'powd') 