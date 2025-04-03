# Requires -Version 5.1
<#
.SYNOPSIS
Shared functionality for PCPow power management scripts

.DESCRIPTION
This module provides common functions and configuration management for PCPow scripts.
It handles process management, user interaction, and power actions with configurable settings.

.NOTES
Version: 1.1.0
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
                    version = "1.1.0"
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
                        "spoolsv", "lsass", "services", "system", "registry", "idle", 
                        "MsMpEng", "fontdrvhost"
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
        [ValidateSet('Info','Success','Warning','Error','Action','Debug')]
        [string]$Level = 'Info'
    )
    
    try {
        # Default to Info color if not found
        $color = if ($script:Config.colors.ContainsKey($Level.ToLower())) {
            $script:Config.colors.$($Level.ToLower())
        } else {
            $script:Config.colors.info
        }
        Write-Host $Message -ForegroundColor $color
    }
    catch {
        # Fallback to basic output if something goes wrong
        Write-Host $Message
    }
}

# Helper function to check if a process is elevated
function Test-ProcessElevated {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int]$ProcessId
    )
    
    try {
        $process = Get-CimInstance Win32_Process -Filter "ProcessId = $ProcessId"
        $owner = Invoke-CimMethod -InputObject $process -MethodName GetOwner
        
        if ($owner) {
            $user = $owner.User
            $domain = $owner.Domain
            
            $adminGroup = (Get-LocalGroup -SID S-1-5-32-544).Name
            $userGroups = Get-LocalGroupMember -Group $adminGroup | 
                Where-Object { $_.Name -eq "$domain\$user" }
            
            return $null -ne $userGroups
        }
    } catch {
        Write-PCPowLog "Failed to check if process is elevated: $_" -Level Debug
        return $false
    }
    return $false
}

# Get the full process tree for the current process
function Get-ProcessTree {
    [CmdletBinding()]
    param()
    
    try {
        # Get current process ID
        $currentPID = $PID
        
        # Get the current process and build tree
        $processTree = @($currentPID)
        
        # Get parent process
        $currentProcess = Get-CimInstance Win32_Process -Filter "ProcessId = $currentPID"
        $parentPID = $currentProcess.ParentProcessId
        $processTree += $parentPID
        
        # Get grandparent process
        $parentProcess = Get-CimInstance Win32_Process -Filter "ProcessId = $parentPID"
        if ($parentProcess) {
            $grandParentPID = $parentProcess.ParentProcessId
            $processTree += $grandParentPID
            
            # Get great-grandparent for deeper terminal hosting scenarios
            $grandParentProcess = Get-CimInstance Win32_Process -Filter "ProcessId = $grandParentPID"
            if ($grandParentProcess) {
                $greatGrandParentPID = $grandParentProcess.ParentProcessId
                $processTree += $greatGrandParentPID
            }
        }
        
        Write-PCPowLog "Process Tree - Current: $currentPID, Parent: $parentPID, GrandParent: $grandParentPID, GreatGrandParent: $($greatGrandParentPID ?? 'none')" -Level Debug
        
        return $processTree
    }
    catch {
        Write-PCPowLog "Failed to get process tree: $_" -Level Warning
        return @($PID) # Return only the current process as fallback
    }
}

# Get the main shell window PID
function Get-ShellWindowPID {
    [CmdletBinding()]
    param()
    
    try {
        Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class ShellHelper {
            [DllImport("user32.dll")]
            public static extern IntPtr GetShellWindow();
            
            [DllImport("user32.dll")]
            public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        }
"@
        
        $shellWindow = [ShellHelper]::GetShellWindow()
        [uint32]$shellPID = 0
        [void][ShellHelper]::GetWindowThreadProcessId($shellWindow, [ref]$shellPID)
        
        Write-PCPowLog "Shell Explorer PID: $shellPID" -Level Debug
        return $shellPID
    }
    catch {
        Write-PCPowLog "Failed to get shell window PID: $_" -Level Warning
        return 0
    }
}

function Get-UserApps {
    [CmdletBinding()]
    param()
    
    try {
        # Get process tree and shell window
        $processTree = Get-ProcessTree
        $shellPID = Get-ShellWindowPID
        
        # Define terminal process types
        $terminalProcesses = @(
            @{Name="powershell"; IsHost=$true},
            @{Name="pwsh"; IsHost=$true},
            @{Name="WindowsTerminal"; IsHost=$true},
            @{Name="OpenConsole"; IsHost=$false},
            @{Name="cmd"; IsHost=$true},
            @{Name="conhost"; IsHost=$false}
        )
        
        # Get all processes
        $allProcesses = Get-Process
        
        # Get terminal host processes
        $terminalHosts = $allProcesses | Where-Object {
            $processName = $_.ProcessName
            ($terminalProcesses | Where-Object { $_.Name -eq $processName -and $_.IsHost }) -and
            -not ($processTree -contains $_.Id)
        }
        
        # Get terminal child processes
        $terminalChildren = @()
        foreach ($terminalHost in $terminalHosts) {
            $children = Get-CimInstance Win32_Process -Filter "ParentProcessId = $($terminalHost.Id)"
            $terminalChildren += $children.ProcessId
        }
        
        # Get all user processes with visible windows
        $userApps = $allProcesses | Where-Object {
            $_.MainWindowHandle -ne 0 -and
            -not ($processTree -contains $_.Id) -and
            $_.Id -ne $shellPID -and
            $_.ProcessName -notin $script:Config.excludedProcesses
        }
        
        # Get all terminal processes regardless of window
        $terminalApps = $allProcesses | Where-Object {
            $process = $_
            (($terminalProcesses | Where-Object { $_.Name -eq $process.ProcessName }) -and
            -not ($processTree -contains $_.Id)) -or
            ($terminalChildren -contains $_.Id)
        }
        
        # Combine and deduplicate
        $combinedApps = $userApps + $terminalApps | Select-Object -Unique
        
        # Add IsElevated property to each process
        $combinedApps | ForEach-Object {
            $_ | Add-Member -NotePropertyName 'IsElevated' -NotePropertyValue (Test-ProcessElevated -ProcessId $_.Id) -Force
        }
        
        return $combinedApps
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
    
    # Get process tree for protection
    $processTree = Get-ProcessTree
    
    # Close File Explorer windows first
    try {
        Write-PCPowLog "Closing File Explorer windows..." -Level Info
        
        # Force close approach for Explorer when in force mode
        if ($useForceMode) {
            Write-PCPowLog "Using force mode for Explorer..." -Level Warning
            try {
                # Try graceful close first for Explorer
                $shell = New-Object -ComObject Shell.Application
                $windows = $shell.Windows()
                foreach ($window in $windows) {
                    try {
                        Write-PCPowLog "Closing Explorer window: $($window.LocationName)" -Level Info
                        $window.Quit()
                        Start-Sleep -Milliseconds 100
                    } catch { }
                }
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
            }
            catch { }
        }
        else {
            # Original graceful approach for Explorer
            try {
                $shell = New-Object -ComObject Shell.Application
                $windows = $shell.Windows()
                foreach ($window in $windows) {
                    try {
                        Write-PCPowLog "Closing Explorer window: $($window.LocationName)" -Level Info
                        $window.Quit()
                        Start-Sleep -Milliseconds 100
                    } catch { }
                }
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
            }
            catch { }
        }

        Start-Sleep -Milliseconds 500 # Give Explorer time to stabilize
    }
    catch {
        Write-PCPowLog "Explorer cleanup encountered issues: $_" -Level Warning
    }
    
    # Get user applications to close
    $userApps = Get-UserApps
    
    if ($userApps.Count -eq 0) {
        Write-PCPowLog "No user applications to close." -Level Info
        return $true
    }
    
    Write-PCPowLog "`nWill close these applications:" -Level Info
    $userApps | Select-Object ProcessName, Id, MainWindowTitle | Format-Table -AutoSize
    
    # Close collected processes
    foreach ($process in $userApps) {
        try {
            # Skip if process is part of our tree
            if ($processTree -contains $process.Id) {
                continue
            }
            
            $isElevated = $process.IsElevated
            $isTerminal = @("powershell", "pwsh", "WindowsTerminal", "OpenConsole", "cmd", "conhost") -contains $process.ProcessName
            
            if ($useForceMode) {
                Write-PCPowLog "Force closing $($process.ProcessName) (PID: $($process.Id))$(if($isElevated){' [Admin]'})" -Level Warning
                if ($isTerminal -and $isElevated) {
                    # For elevated terminals, try to close gracefully first
                    $process.CloseMainWindow() | Out-Null
                    Start-Sleep -Milliseconds 500
                    if (-not $process.HasExited) {
                        $process | Stop-Process -Force -ErrorAction SilentlyContinue
                    }
                } else {
                    $process | Stop-Process -Force -ErrorAction SilentlyContinue
                }
            } else {
                Write-PCPowLog "Gracefully closing $($process.ProcessName) (PID: $($process.Id))$(if($isElevated){' [Admin]'})" -Level Info
                if ($process.MainWindowHandle -ne 0) {
                    $process.CloseMainWindow() | Out-Null
                    Start-Sleep -Milliseconds 500
                }
                
                if (-not $process.HasExited) {
                    Write-PCPowLog "Forcing close of $($process.ProcessName)" -Level Warning
                    $process | Stop-Process -Force -ErrorAction SilentlyContinue
                }
            }
        } catch {
            Write-PCPowLog "Failed to close $($process.ProcessName): $_" -Level Warning
        }
    }
    
    # Final check for suspended processes
    $suspendedProcesses = Get-Process | Where-Object { 
        $_.Responding -eq $false -and 
        $_.ProcessName -notin $script:Config.excludedProcesses -and
        -not ($processTree -contains $_.Id)
    }
    
    if ($suspendedProcesses) {
        Write-PCPowLog "`nAttempting to close suspended processes..." -Level Warning
        foreach ($process in $suspendedProcesses) {
            try {
                Write-PCPowLog "Closing suspended $($process.ProcessName) (PID: $($process.Id))" -Level Warning
                $process | Stop-Process -Force -ErrorAction SilentlyContinue
            } catch {
                Write-PCPowLog "Failed to close suspended process: $_" -Level Warning
            }
        }
    }

    # Final verification and cleanup
    $remainingProcesses = $userApps | Where-Object { -not $_.HasExited }
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
    $remainingCount = ($userApps | Where-Object { -not $_.HasExited }).Count
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
        Write-PCPowLog "Initiating $Action sequence..." -Level Action
        
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
        if (-not (Close-Applications -TimeoutMS $timeout -NoGraceful:$script:Config.NoGraceful)) {
            if (-not $useForce) {
                throw "Some applications could not be closed cleanly"
            }
        }

        Write-PCPowLog "Performing $Action..." -Level Action
        
        # Ensure we're not in a nested PowerShell session
        $isNested = [bool]$PSCommandPath

        switch ($Action) {
            'Restart' { 
                try {
                    if ($useForce) {
                        Write-PCPowLog "Forcing restart..." -Level Warning
                        try {
                            if ($isNested) {
                                & shutdown.exe /r /f /t 0
                            } else {
                                Start-Process "shutdown.exe" -ArgumentList "/r /f /t 0" -NoNewWindow
                            }
                        } catch {
                            # Fallback to Restart-Computer if shutdown.exe fails
                            Restart-Computer -Force -ErrorAction Stop
                        }
                    } else {
                        Write-PCPowLog "Initiating graceful restart..." -Level Info
                        try {
                            if ($isNested) {
                                & shutdown.exe /r /t 10
                            } else {
                                Start-Process "shutdown.exe" -ArgumentList "/r /t 10" -NoNewWindow
                            }
                        } catch {
                            # Fallback to Restart-Computer if shutdown.exe fails
                            Restart-Computer -ErrorAction Stop
                        }
                    }
                    
                    # Give the restart command time to take effect
                    Start-Sleep -Seconds 2
                }
                catch {
                    Write-PCPowLog "Failed to initiate restart: $_" -Level Error
                    throw
                }
            }
            'Shutdown' { 
                try {
                    if ($useForce) {
                        Write-PCPowLog "Forcing shutdown..." -Level Warning
                        # Try multiple methods for reliability
                        try {
                            if ($isNested) {
                                & shutdown.exe /s /f /t 0
                            } else {
                                Start-Process "shutdown.exe" -ArgumentList "/s /f /t 0" -NoNewWindow
                            }
                        } catch {
                            # Fallback to Stop-Computer if shutdown.exe fails
                            Stop-Computer -Force -ErrorAction Stop
                        }
                    } else {
                        Write-PCPowLog "Initiating graceful shutdown..." -Level Info
                        # Give apps a bit more time to close in graceful mode
                        try {
                            if ($isNested) {
                                & shutdown.exe /s /t 10
                            } else {
                                Start-Process "shutdown.exe" -ArgumentList "/s /t 10" -NoNewWindow
                            }
                        } catch {
                            # Fallback to Stop-Computer if shutdown.exe fails
                            Stop-Computer -ErrorAction Stop
                        }
                    }
                    
                    # Give the shutdown command time to take effect
                    Start-Sleep -Seconds 2
                }
                catch {
                    Write-PCPowLog "Failed to initiate shutdown: $_" -Level Error
                    throw
                }
            }
            'Sleep' { 
                try {
                    Write-PCPowLog "Preparing system for sleep..." -Level Info
                    
                    # Check for sleep blockers
                    Write-PCPowLog "Checking for sleep prevention requests..." -Level Info
                    $requests = & powercfg /requests
                    if ($requests -match "None.") {
                        Write-PCPowLog "No active sleep prevention requests." -Level Success
                    } else {
                        Write-PCPowLog "Active sleep prevention requests found:" -Level Warning
                        $requests | ForEach-Object { Write-PCPowLog $_ -Level Warning }
                    }
                    
                    # Clear Windows Update operations
                    Write-PCPowLog "Resetting Windows Update service..." -Level Info
                    Stop-Service -Name "wuauserv" -Force -ErrorAction SilentlyContinue
                    Start-Service -Name "wuauserv" -ErrorAction SilentlyContinue
                    
                    Add-Type -TypeDefinition @"
                    using System;
                    using System.Runtime.InteropServices;
                    
                    public class SleepManager {
                        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
                        public static extern uint SetThreadExecutionState(uint esFlags);
                        
                        [DllImport("PowrProf.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
                        public static extern bool SetSuspendState(bool hibernate, bool forceCritical, bool disableWakeEvent);
                        
                        public static void ClearWakeRequirements() {
                            // ES_CONTINUOUS = 0x80000000
                            SetThreadExecutionState(0x80000000);
                        }
                        
                        public static bool ForceSleep() {
                            ClearWakeRequirements();
                            return SetSuspendState(false, true, true);
                        }
                    }
"@
                    
                    Write-PCPowLog "Clearing system wake locks..." -Level Info
                    [SleepManager]::ClearWakeRequirements()
                    
                    # Try to prevent system from waking up immediately
                    $preventWakeHandle = $null
                    try {
                        $preventWakeHandle = [System.Threading.Mutex]::new($true, "PCPowPreventWake")
                    } catch { }
                    
                    Write-PCPowLog "Initiating sleep sequence..." -Level Action
                    Start-Sleep -Seconds 2
                    
                    # Try sleep methods in sequence
                    try {
                        # Method 1: Direct API call
                        Write-PCPowLog "Attempting sleep via API..." -Level Info
                        if ([SleepManager]::ForceSleep()) {
                            exit 0
                        }
                        
                        Start-Sleep -Seconds 3
                        
                        # Method 2: PowerProf.dll
                        Write-PCPowLog "Attempting sleep via PowerProf.dll..." -Level Info
                        rundll32.exe powrprof.dll,SetSuspendState 0,1,0
                        
                        Start-Sleep -Seconds 3
                        
                        # Method 3: Hibernate fallback
                        Write-PCPowLog "Attempting hibernate as fallback..." -Level Info
                        shutdown.exe /h
                        
                        Start-Sleep -Seconds 3
                        
                        # Final attempt with powercfg reset
                        Write-PCPowLog "Final sleep attempt with powercfg reset..." -Level Warning
                        powercfg -hibernate off
                        powercfg -hibernate on
                        Start-Sleep -Seconds 1
                        [SleepManager]::ForceSleep()
                    }
                    catch {
                        Write-PCPowLog "Failed to initiate sleep: $_" -Level Error
                        throw
                    }
                    finally {
                        if ($preventWakeHandle) {
                            $preventWakeHandle.ReleaseMutex()
                            $preventWakeHandle.Dispose()
                        }
                    }
                }
                catch {
                    Write-PCPowLog "Failed to initiate sleep: $_" -Level Error
                    throw
                }
            }
        }
    }
    catch {
        Write-PCPowLog "Error during $Action operation: $_" -Level Error
        throw
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

# Define the functions with approved verbs
function Start-PCSleep { param([switch]$Force) Invoke-PowerAction -Action Sleep -Force:$Force }
function Restart-PC { param([switch]$Force) Invoke-PowerAction -Action Restart -Force:$Force }
function Stop-PC { param([switch]$Force) Invoke-PowerAction -Action Shutdown -Force:$Force }

# Export module members
Export-ModuleMember -Function @(
    'Initialize-PCPowConfig',
    'Get-UserApps',
    'Start-PCSleep',
    'Restart-PC',
    'Stop-PC',
    'Close-Applications',
    'Show-ConfirmationPrompt',
    'Invoke-PowerAction',
    'Write-PCPowLog',
    'Test-ProcessElevated',
    'Get-ProcessTree'
) -Alias @('pows', 'powr', 'powd')

# Set up aliases
New-Alias -Name pows -Value Start-PCSleep -Force
New-Alias -Name powr -Value Restart-PC -Force
New-Alias -Name powd -Value Stop-PC -Force