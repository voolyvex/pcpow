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
        # First try module directory
        if (Test-Path $ConfigPath) {
            $script:Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
            Write-PCPowLog "Loaded config from module path: $ConfigPath" -Level Info
        } 
        # Then try WindowsApps directory
        else {
            $windowsAppsConfig = Join-Path $env:USERPROFILE "AppData\Local\Microsoft\WindowsApps\pcpow.config.json"
            if (Test-Path $windowsAppsConfig) {
                $script:Config = Get-Content $windowsAppsConfig -Raw | ConvertFrom-Json
                Write-PCPowLog "Loaded config from WindowsApps: $windowsAppsConfig" -Level Info
            }
            else {
                Write-Warning "Configuration file not found at $ConfigPath or $windowsAppsConfig. Using default settings."
                $script:Config = @{
                    version = "1.0.0"
                    timeoutMS = 5000
                    AlwaysForce = $false
                    NoGraceful = $false
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

        # Ensure all required properties exist with proper error checking
        if (-not (Get-Member -InputObject $script:Config -Name 'AlwaysForce')) {
            $script:Config | Add-Member -NotePropertyName 'AlwaysForce' -NotePropertyValue $false
        }
        if (-not (Get-Member -InputObject $script:Config -Name 'NoGraceful')) {
            $script:Config | Add-Member -NotePropertyName 'NoGraceful' -NotePropertyValue $false
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
    param(
        [int]$TimeoutMS = 5000,
        [switch]$NoGraceful
    )
    
    # Start timeout tracking immediately
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    # Determine if we should skip graceful shutdown
    $useForceMode = $script:Config.AlwaysForce -or $NoGraceful
    
    # Track processes that need force closing - use HashSet for better performance
    $remainingProcesses = [System.Collections.Generic.HashSet[string]]::new()
    
    # Optimize process collection - do it once and filter efficiently
    $allProcesses = Get-Process
    $currentSession = ($allProcesses | Where-Object { $_.Id -eq $PID }).SessionId
    
    # More efficient process filtering
    $processes = $allProcesses | Where-Object {
        ($_.MainWindowHandle -ne 0 -or $_.ProcessName -eq 'explorer') -and 
        $_.SessionId -eq $currentSession -and
        $_.ProcessName -notin $script:Config.excludedProcesses
    } | Sort-Object -Property WorkingSet64 -Descending

    # Categorize processes by importance
    $criticalProcesses = $processes | Where-Object {
        $_.ProcessName -match "^(explorer|SearchUI|StartMenuExperienceHost|Cortana|SearchApp)$"
    }
    $userProcesses = $processes | Where-Object {
        $_ -notin $criticalProcesses
    }

    # Handle File Explorer windows first - optimized for speed
    try {
        Write-PCPowLog "Closing File Explorer windows..." -Level Info
        
        # Force close approach for Explorer when in force mode
        if ($useForceMode) {
            Write-PCPowLog "Using force mode for Explorer..." -Level Warning
            try {
                # Try graceful close first for Explorer
                $shell = New-Object -ComObject Shell.Application
                $shell.Windows() | ForEach-Object {
                    try {
                        if ($_.FullName -match "explorer\.exe$") { 
                            $_.Quit()
                            Start-Sleep -Milliseconds 100
                        }
                    } catch { }
                }
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null

                # Only force close if graceful failed
                Get-Process explorer -ErrorAction SilentlyContinue | ForEach-Object {
                    if ($_.MainWindowHandle -ne 0) {
                        $_.CloseMainWindow() | Out-Null
                        Start-Sleep -Milliseconds 100
                        if (-not $_.HasExited) {
                            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
                        }
                    }
                }
                Start-Sleep -Milliseconds 500
                Start-Process explorer
            }
            catch { }
        }
        else {
            # Original graceful approach for Explorer
            try {
                $shell = New-Object -ComObject Shell.Application
                $shell.Windows() | ForEach-Object {
                    try {
                        if ($_.FullName -match "explorer\.exe$") { 
                            $_.Quit()
                            Start-Sleep -Milliseconds 100
                        }
                    } catch { }
                }
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
            }
            catch { }
        }

        # Always clean up Explorer state
        try {
            $regPaths = @(
                "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\BrowseForFolder",
                "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\LastVisitedPidlMRU",
                "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\LastVisitedPidlMRULegacy"
            )
            foreach ($path in $regPaths) {
                if (Test-Path $path) {
                    Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }
        catch { }

        Start-Sleep -Milliseconds 500 # Give Explorer time to stabilize
    }
    catch {
        Write-PCPowLog "Explorer cleanup encountered issues: $_" -Level Warning
    }

    # Handle user processes first (aggressive)
    if ($useForceMode) {
        Write-PCPowLog "Force closing user applications..." -Level Warning
        foreach ($process in $userProcesses) {
            try {
                Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
                Start-Sleep -Milliseconds 50
            } catch { }
        }
    }
    else {
        # Original graceful shutdown logic for user processes
        foreach ($process in $userProcesses) {
            if ($stopwatch.ElapsedMilliseconds -ge $TimeoutMS) {
                break
            }
            try {
                if (-not $process.HasExited) {
                    Write-PCPowLog "Closing $($process.Name) (PID: $($process.Id))..." -Level Info
                    if ($process.CloseMainWindow()) {
                        $process.WaitForExit(2000)
                    }
                }
            } catch { }
        }
    }

    # Handle critical processes more carefully
    foreach ($process in $criticalProcesses) {
        try {
            if (-not $process.HasExited) {
                Write-PCPowLog "Carefully closing $($process.Name)..." -Level Info
                if ($process.CloseMainWindow()) {
                    $process.WaitForExit(2000)
                }
                if (-not $process.HasExited) {
                    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
                }
            }
        } catch { }
    }

    # Final verification and cleanup
    $remainingProcesses = $processes | Where-Object { -not $_.HasExited }
    if ($remainingProcesses) {
        Write-PCPowLog "Force closing remaining applications..." -Level Warning
        foreach ($process in $remainingProcesses) {
            try {
                taskkill /F /IM "$($process.ProcessName).exe" /T 2>&1 | Out-Null
            } catch { }
        }
    }

    # Final cleanup
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # Verify all processes are closed
    $remainingCount = ($processes | Where-Object { -not $_.HasExited }).Count
    return $remainingCount -eq 0
}

function Show-ConfirmationPrompt {
    [CmdletBinding()]
    param(
        [string]$ActionType
    )

    # If AlwaysForce is enabled, skip confirmation
    if ($script:Config.AlwaysForce) {
        Write-PCPowLog "Skipping confirmation (AlwaysForce mode)" -Level Info
        return $true
    }

    $modeInfo = if ($script:Config.NoGraceful) { " (NoGraceful mode)" } else { "" }
    $title = "Confirm $ActionType$modeInfo"
    $message = "WARNING: This will close all running programs and $ActionType.`nSave any important work before continuing.`nProceed?"

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Confirm $ActionType"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Cancel operation"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $message, $options, 1)
    
    # If user selected No, return false
    if ($result -ne 0) {
        Write-PCPowLog "Operation cancelled by user" -Level Warning
        return $false
    }
    return $true
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
        
        # Check confirmation first
        if (-not $Force -and -not $script:Config.AlwaysForce) {
            if (-not (Show-ConfirmationPrompt -ActionType $Action)) {
                Write-PCPowLog "Operation cancelled" -Level Warning
                return
            }
        }
        
        # Use configuration settings for Force and NoGraceful
        $useForce = $Force -or $script:Config.AlwaysForce
        
        # Determine timeout based on force mode
        $timeout = if ($useForce) { 3000 } else { $script:Config.TimeoutMS }
        
        # Close applications with optimized parameters
        if (-not (Close-Applications -TimeoutMS $timeout -NoGraceful $script:Config.NoGraceful)) {
            if (-not $useForce) {
                throw "Some applications could not be closed cleanly"
            }
        }

        Write-Host "Performing $Action..." -ForegroundColor $script:Config.Colors.Action
        switch ($Action) {
            'Restart' { 
                if ($useForce) {
                    shutdown /r /f /t 0
                } else {
                    Restart-Computer -Force:$useForce
                }
            }
            'Shutdown' { 
                if ($useForce) {
                    shutdown /s /f /t 0
                } else {
                    Stop-Computer -Force:$useForce
                }
            }
            'Sleep' { 
                Add-Type -AssemblyName System.Windows.Forms
                [System.Windows.Forms.Application]::SetSuspendState("Suspend", $useForce, $false)
            }
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