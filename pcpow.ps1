# Requires -Version 5.1
<#
.SYNOPSIS
Windows Power Management Utility (v1.2.2)

.DESCRIPTION
Safely closes applications and performs power actions (sleep, restart, shutdown).
Includes Wake-on-LAN functionality via pcpow.bat.

.PARAMETER Action
The power action to perform: Sleep, Restart, or Shutdown. Required unless -FormatMAC is used.

.PARAMETER Force
If specified, forces applications to close without waiting for them to exit gracefully.

.PARAMETER SkipAction
If specified, only closes applications without performing the power action (for testing).

.PARAMETER FormatMAC
Internal use by pcpow.bat. Takes a MAC address string and returns it in colon-separated format after validation.

.EXAMPLE
.\PCPow.ps1 -Action Sleep
Safely closes applications and puts the computer to sleep.

.EXAMPLE
.\PCPow.ps1 -Action Restart -Force
Forces applications to close and restarts the computer.

.EXAMPLE
.\PCPow.ps1 -FormatMAC "01-23-45-67-89-AB"
Outputs: 01:23:45:67:89:ab

.NOTES
Implements Shadow Worker Enhanced Autonomy Protocol v2.0
Path verification: Section 2.1
Terminal standards: Section 3.2
Requires PowerShell 5.1 or later.
Run setup-shortcuts.ps1 for easy command-line access.
Use update-profile.ps1 if PowerShell profile issues occur.
#>

[CmdletBinding(DefaultParameterSetName='PowerAction')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='PowerAction')]
    [ValidateSet('Sleep', 'Restart', 'Shutdown')]
    [string]$Action,

    [Parameter(ParameterSetName='PowerAction')]
    [switch]$Force,

    [Parameter(ParameterSetName='PowerAction')]
    [switch]$SkipAction,

    [Parameter(Mandatory=$true, ParameterSetName='FormatMAC')]
    [string]$FormatMAC
)

#region Helper Functions

function Write-LogMessage {
    param (
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Action', 'Debug')]
        [string]$Level = 'Info'
    )

    # Load colors from config if available, otherwise use defaults
    $logColors = $script:config.colors # Use script scope config if loaded
    if (-not $logColors) {
        $logColors = @{
            warning = "Yellow"
            success = "Green"
            error = "Red"
            info = "Cyan"
            action = "Magenta"
            debug = "DarkGray"
        }
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level.ToLower()) {
        'info'    { if ($logColors.info)    { $logColors.info    } else { 'Cyan' } }
        'warning' { if ($logColors.warning) { $logColors.warning } else { 'Yellow' } }
        'error'   { if ($logColors.error)   { $logColors.error   } else { 'Red' } }
        'success' { if ($logColors.success) { $logColors.success } else { 'Green' } }
        'action'  { if ($logColors.action)  { $logColors.action  } else { 'Magenta'} }
        'debug'   { if ($logColors.debug)   { $logColors.debug   } else { 'DarkGray' } }
        default   { 'White' }
    }

    # Console output
    Write-Host "[$timestamp] " -NoNewline -ForegroundColor DarkGray
    Write-Host $Message -ForegroundColor $color

    # Append to log file
    try {
        $logDir = Join-Path $env:USERPROFILE "AppData\Local\PCPow\logs"
        $logFile = Join-Path $logDir "pcpow_$(Get-Date -Format 'yyyyMMdd').log"
        
        # Create log directory if it doesn't exist
        if (-not (Test-Path -Path $logDir -PathType Container)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        
        # Format log entry
        $logEntry = "[$timestamp] [$Level] $Message"
        Add-Content -Path $logFile -Value $logEntry -Encoding UTF8
    }
    catch {
        # Silent fail for logging to file - don't disrupt main operations if logging fails
        Write-Host "Failed to write to log file: $_" -ForegroundColor $logColors.error
    }
}

# Consistent path testing helper
function Test-PathExists {
    param (
        [string]$Path,
        [ValidateSet('Any', 'Leaf', 'Container')]
        [string]$PathType = 'Any' # Leaf=File, Container=Directory
    )
    process {
        if ([string]::IsNullOrWhiteSpace($Path)) {
            Write-LogMessage "Test-PathExists: Received null or empty path." -Level Debug
            return $false
        }
        try {
            # Resolve relative paths for robustness
            $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            if (-not (Test-Path -LiteralPath $resolvedPath -IsValid)) {
                 Write-LogMessage "Test-PathExists: Path '$resolvedPath' is invalid." -Level Debug
                 return $false
            }

            $exists = Test-Path -LiteralPath $resolvedPath
            if (-not $exists) {
                Write-LogMessage "Test-PathExists: Path '$resolvedPath' does not exist." -Level Debug
                return $false
            }

            if ($PathType -eq 'Any') {
                return $true
            }

            $item = Get-Item -LiteralPath $resolvedPath -ErrorAction SilentlyContinue
            if (-not $item) {
                 Write-LogMessage "Test-PathExists: Could not get item for path '$resolvedPath'." -Level Debug
                 return $false # Should exist based on Test-Path, but check anyway
            }

            if ($PathType -eq 'Leaf' -and -not $item.PSIsContainer) {
                return $true
            }
            if ($PathType -eq 'Container' -and $item.PSIsContainer) {
                return $true
            }

             Write-LogMessage "Test-PathExists: Path '$resolvedPath' exists but type mismatch (Expected: $PathType)." -Level Debug
             return $false
        } catch {
            Write-LogMessage "Test-PathExists: Error testing path '$Path': $_" -Level Warning
            return $false
        }
    }
}


function Get-ConfigurationSettings {
    # Determine script root reliably
    if ($PSScriptRoot) {
        $scriptRoot = $PSScriptRoot
    } else {
        # Fallback for environments where $PSScriptRoot might not be set (e.g., ISE)
        try {
            $scriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent -ErrorAction Stop
        } catch {
            Write-LogMessage "Critical Error: Could not determine script root directory." -Level Error
            throw "Failed to determine script root directory."
        }
    }

    # Verify script root exists
    if (-not (Test-PathExists -Path $scriptRoot -PathType Container)) {
         Write-LogMessage "Critical Error: Script root directory '$scriptRoot' not found or invalid." -Level Error
         throw "Script root directory '$scriptRoot' not found or invalid."
    }

    $configPath = Join-Path $scriptRoot "pcpow.config.json"
    $configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($configPath) # Normalize

    # Default settings
    $defaultConfig = @{
        version = "1.2.1" # Track config version
        timeoutMS = 5000
        AlwaysForce = $false
        NoGraceful = $false
        colors = @{ # Default color scheme
            warning = "Yellow"
            success = "Green"
            error = "Red"
            info = "Cyan"
            action = "Magenta"
            debug = "DarkGray"
        }
        excludedProcesses = @(
            "svchost", "csrss", "smss", "wininit", "winlogon",
            "spoolsv", "lsass", "services", "system", "registry", "idle",
            "dwm", "fontdrvhost", "Memory Compression", "SearchIndexer",
            "ShellExperienceHost", "sihost", "taskmgr", "SecurityHealthService",
            "RuntimeBroker", "dllhost", "ctfmon", "conhost", "WmiPrvSE",
            "MsMpEng", "SearchHost", "SearchApp", "StartMenuExperienceHost",
            "ApplicationFrameHost", "explorer", "powershell", "powershell_ise",
            "WindowsTerminal", "cmd", "pwsh",
            "NVDisplay.Container", "TextInputHost", "SystemSettings" # Added some common persistent ones
        )
        # Wake-on-LAN settings are handled separately by Setup-WakeOnLAN.ps1
    }

    # Load config if it exists
    if (Test-PathExists -Path $configPath -PathType Leaf) {
        try {
            $configJson = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 -ErrorAction Stop
            $loadedConfig = $configJson | ConvertFrom-Json -ErrorAction Stop

            # Merge loaded config with defaults (prioritize loaded values)
            $mergedConfig = $defaultConfig.Clone() # Start with defaults
            foreach ($key in $loadedConfig.PSObject.Properties.Name) {
                if ($mergedConfig.ContainsKey($key)) {
                    # Handle nested objects like 'colors'
                    if ($mergedConfig[$key] -is [hashtable] -and $loadedConfig[$key] -is [hashtable]) {
                        foreach ($nestedKey in $loadedConfig[$key].Keys) {
                             if ($mergedConfig[$key].ContainsKey($nestedKey)) {
                                 $mergedConfig[$key][$nestedKey] = $loadedConfig[$key][$nestedKey]
                             } else {
                                 Write-LogMessage "Ignoring unknown nested key '$nestedKey' in config section '$key'." -Level Debug
                             }
                        }
                    } else {
                       $mergedConfig[$key] = $loadedConfig[$key]
                    }
                } else {
                     Write-LogMessage "Ignoring unknown key '$key' in config file." -Level Debug
                }
            }
            $defaultConfig = $mergedConfig # Use the merged result
            Write-LogMessage "Loaded configuration from $configPath" -Level Info

        } catch {
            Write-LogMessage "Error reading or parsing config file '$configPath': $_. Using default configuration." -Level Warning
        }
    } else {
        Write-LogMessage "Configuration file not found: $configPath. Using default configuration." -Level Warning

        # Attempt to create a new config file with defaults
        try {
            $configDir = Split-Path -Parent $configPath
            if (-not (Test-PathExists -Path $configDir -PathType Container)) {
                New-Item -ItemType Directory -Path $configDir -Force -ErrorAction Stop | Out-Null
            }
            $defaultConfigJson = $defaultConfig | ConvertTo-Json -Depth 5
            Set-Content -LiteralPath $configPath -Value $defaultConfigJson -Encoding UTF8 -ErrorAction Stop
            Write-LogMessage "Created new configuration file with defaults: $configPath" -Level Success
        } catch {
            Write-LogMessage "Failed to create new configuration file '$configPath': $_" -Level Warning
        }
    }

    # --- Configuration Validation ---
    if ($defaultConfig.TimeoutMS -isnot [int] -or $defaultConfig.TimeoutMS -lt 1000) {
        Write-LogMessage "Invalid TimeoutMS value ('$($defaultConfig.TimeoutMS)') in config, must be an integer >= 1000. Using default 5000." -Level Warning
        $defaultConfig.TimeoutMS = 5000
    }
    if ($defaultConfig.AlwaysForce -isnot [bool]) {
         Write-LogMessage "Invalid AlwaysForce value ('$($defaultConfig.AlwaysForce)') in config, must be true or false. Using default false." -Level Warning
         $defaultConfig.AlwaysForce = $false
    }
    if ($defaultConfig.NoGraceful -isnot [bool]) {
         Write-LogMessage "Invalid NoGraceful value ('$($defaultConfig.NoGraceful)') in config, must be true or false. Using default false." -Level Warning
         $defaultConfig.NoGraceful = $false
    }
    if ($defaultConfig.excludedProcesses -isnot [array]) {
        Write-LogMessage "Invalid excludedProcesses format in config, must be an array. Using default list." -Level Warning
        # Reset to default list (already set, just log)
    } else {
        # Ensure core protected processes are always present
        $coreProtected = @("explorer", "powershell", "pwsh", "cmd", "WindowsTerminal")
        foreach ($proc in $coreProtected) {
            if ($defaultConfig.excludedProcesses -notcontains $proc) {
                 Write-LogMessage "Ensuring core process '$proc' is in excludedProcesses list." -Level Debug
                 $defaultConfig.excludedProcesses += $proc
            }
        }
        # Convert all to lowercase for case-insensitive matching later
        $defaultConfig.excludedProcesses = $defaultConfig.excludedProcesses | ForEach-Object { $_.ToLowerInvariant() }
    }
     # Validate colors (ensure they are valid ConsoleColor names)
     if ($defaultConfig.colors -is [hashtable]) {
        $validColors = [System.Enum]::GetNames([System.ConsoleColor])
        foreach ($key in $defaultConfig.colors.Keys) {
            if ($defaultConfig.colors[$key] -notin $validColors) {
                Write-LogMessage "Invalid color value '$($defaultConfig.colors[$key])' for key '$key' in config. Using default for this color." -Level Warning
                # Let Write-LogMessage handle the default color fallback
                $defaultConfig.colors.Remove($key) # Remove invalid key to trigger default
            }
        }
     } else {
         Write-LogMessage "Invalid 'colors' section in config, must be a hashtable. Using default colors." -Level Warning
         # Reset to default colors (already set, just log)
     }


    # Assign to script scope for global access within this script run
    $script:config = $defaultConfig
    return $script:config
}

# Robustly close user applications
function Close-UserProcesses {
    param(
        [switch]$ForceParam, # Renamed to avoid conflict with $Force variable used later
        [int]$TimeoutMSParam,
        [array]$ExcludedProcessesParam,
        [ref]$ErrorState # Optional reference variable to report errors back
    )

    $closeSuccess = $true # Assume success initially
    $failedProcesses = @()
    $startTime = Get-Date

    # Create a log header with detailed information about the session
    Write-LogMessage "============ CLOSE PROCESSES SESSION START ============" -Level Info
    Write-LogMessage "Date/Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level Info
    Write-LogMessage "User: $($env:USERNAME) | Machine: $($env:COMPUTERNAME)" -Level Info 
    Write-LogMessage "Force: $($ForceParam.IsPresent) | TimeoutMS: $TimeoutMSParam" -Level Info
    Write-LogMessage "PSVersion: $($PSVersionTable.PSVersion)" -Level Info
    Write-LogMessage "Working Directory: $(Get-Location)" -Level Info

    try {
        # --- Process Tree Protection ---
        Write-LogMessage "Identifying protected process tree..." -Level Debug
        $currentPID = $PID
        $parentProcessIDs = @($currentPID) # Start with self
        $maxDepth = 4 # Protect self, parent, grandparent, great-grandparent
        $currentParentPID = $currentPID

        for ($i = 0; $i -lt $maxDepth; $i++) {
            try {
                # Use CIM for better error handling and performance than Get-Process
                $processInfo = Get-CimInstance Win32_Process -Filter "ProcessId = $currentParentPID" -ErrorAction Stop
                if (-not $processInfo) { break } # Should not happen but check

                $parentPID = $processInfo.ParentProcessId
                if ($parentPID -eq 0 -or $parentPID -in $parentProcessIDs) { break } # Reached root or loop

                $parentProcessIDs += $parentPID
                $currentParentPID = $parentPID
            } catch {
                Write-LogMessage "Warning: Could not get parent process info for PID $currentParentPID - $_" -Level Warning                break # Stop climbing if error occurs
            }
        }
        Write-LogMessage "Protected process ancestors (PID): $($parentProcessIDs -join ', ')" -Level Debug

        # Also protect direct children of our ancestors to cover complex hosting (e.g., Terminal tabs)
        $childProcessesToProtect = @()
        foreach ($ancestorPID in $parentProcessIDs) {
            try {
                $children = Get-CimInstance Win32_Process -Filter "ParentProcessId = $ancestorPID" -ErrorAction SilentlyContinue
                if ($children) {
                    $childPIDs = $children | Select-Object -ExpandProperty ProcessId
                    $childProcessesToProtect += $childPIDs
                }
            } catch {
                Write-LogMessage "Warning: Could not get child processes for PID $ancestorPID - $_" -Level Warning
            }
        }
        $parentProcessIDs += $childProcessesToProtect | Sort-Object -Unique
        Write-LogMessage "Full protected process tree (PID): $($parentProcessIDs -join ', ')" -Level Info

        # --- End Process Tree Protection ---

        # --- Close File Explorer Windows ---
        Write-LogMessage "Closing File Explorer windows..." -Level Info
        $shell = $null
        try {
            $shell = New-Object -ComObject Shell.Application -ErrorAction Stop
            $windows = $shell.Windows()
            if ($windows.Count -gt 0) {
                # Iterate backwards as closing windows modifies the collection
                for ($i = $windows.Count - 1; $i -ge 0; $i--) {
                    $window = $windows.Item($i)
                    if ($null -eq $window) { continue } # Skip if window disappeared

                    $windowName = "Unknown"
                    try { $windowName = $window.LocationName } catch {} # Best effort name

                    # Check if it's actually a File Explorer window (heuristic)
                    $isExplorer = $false
                    try {
                        $app = $window.Application
                        if ($null -ne $app -and $app.GetType().Name -eq 'IWebBrowser2') {
                           # Further check if path is a directory or known explorer CLSID
                           $locationUrl = $window.LocationURL
                           if ($locationUrl -like "file:///*" -or $locationUrl -like "::{*}*") {
                               $isExplorer = $true
                           }
                        }
                    } catch {}

                    if ($isExplorer) {
                        Write-LogMessage "Attempting to close Explorer window: '$windowName'" -Level Debug
                        try {
                            $window.Quit()
                            Start-Sleep -Milliseconds 100 # Small delay
                        } catch {
                            Write-LogMessage "Warning: Failed to gracefully close Explorer window '$windowName': ${_}" -Level Warning
                            # Explorer process itself is handled by exclusion list
                        }
                    }
                }
            } else {
                 Write-LogMessage "No File Explorer windows found open." -Level Debug
            }
        } catch {
             Write-LogMessage "Error interacting with Shell.Application COM object: $($_)" -Level Warning
             # Continue without Explorer closing if COM fails
        } finally {
            if ($null -ne $shell) {
                try {
                    while ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) -gt 0) {}
                    $shell = $null
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                     Write-LogMessage "Released Shell.Application COM object." -Level Debug
                } catch {
                    Write-LogMessage "Warning: Error releasing Shell.Application COM object: $($_)" -Level Warning
                }
            }
        }
        # --- End Close File Explorer Windows ---

        # --- Get User Applications to Close ---
        Write-LogMessage "Scanning for user applications to close..." -Level Info
        $userProcessesToClose = @()
        
        # Get current session ID
        $currentUserSessionId = (Get-Process -Id $PID).SessionId
        Write-LogMessage "Current user session ID: $currentUserSessionId" -Level Debug
        
        # Log exclusion list
        Write-LogMessage "Excluded Processes: $($ExcludedProcessesParam -join ', ')" -Level Debug
        
        try {
            # First attempt - standard user processes with windows and active sessions
            $candidateProcesses = Get-Process | Where-Object {
                $_.SessionId -eq $currentUserSessionId -and
                $_.MainWindowHandle -ne [System.IntPtr]::Zero -and # Has a window
                $_.Id -notin $parentProcessIDs -and               # Not in protected tree
                $_.ProcessName.ToLowerInvariant() -notin $ExcludedProcessesParam # Not in exclusion list (already lowercase)
            } | Sort-Object -Property ProcessName
            
            Write-LogMessage "Found $($candidateProcesses.Count) user processes with visible windows" -Level Info
            
            # Check for specifically known IDE and app processes that might not have visible windows but should be closed
            $knownAppsToAlwaysClose = @(
                'Code',           # VS Code
                'cursor',         # Cursor IDE
                'Cursor',         # Cursor IDE different casing
                'CursorUpdater',  # Cursor IDE updater
                'devenv',         # Visual Studio
                'rider64',        # JetBrains Rider 
                'idea64',         # IntelliJ IDEA
                'pycharm64',      # PyCharm
                'webstorm64',     # WebStorm
                'phpstorm64',     # PHPStorm
                'clion64',        # CLion
                'goland64',       # GoLand
                'msedge',         # Microsoft Edge
                'chrome',         # Google Chrome
                'firefox',        # Firefox
                'slack',          # Slack
                'teams',          # Microsoft Teams
                'outlook',        # Microsoft Outlook
                'winword',        # Microsoft Word
                'excel',          # Microsoft Excel
                'powerpnt',       # Microsoft PowerPoint
                'onenote',        # Microsoft OneNote
                'AcroRd32',       # Adobe Reader
                'Acrobat',        # Adobe Acrobat
                'Spotify',        # Spotify
                'Discord',        # Discord
                'Zoom',           # Zoom
                'vlc',            # VLC
                'mpc-hc',         # Media Player Classic
                'explorer'        # Explorer (we'll handle special case)
            )
            
            # Explicitly look for these applications regardless of window state
            $knownAppProcesses = @()
            foreach ($appName in $knownAppsToAlwaysClose) {
                # Skip explorer as it's handled separately and is typically in the exclude list
                if ($appName -eq 'explorer' -and $ExcludedProcessesParam -contains 'explorer') {
                    continue
                }
                
                $processes = Get-Process -Name $appName -ErrorAction SilentlyContinue | Where-Object {
                    $_.SessionId -eq $currentUserSessionId -and
                    $_.Id -notin $parentProcessIDs -and
                    $_.ProcessName.ToLowerInvariant() -notin $ExcludedProcessesParam
                }
                if ($processes) {
                    $knownAppProcesses += $processes
                    Write-LogMessage "Found specific app process: $appName (Count: $($processes.Count))" -Level Debug
                }
            }
            
            # Combine the regular window processes with known app processes, filtering duplicates
            $allCandidateProcesses = @($candidateProcesses) + @($knownAppProcesses) | Sort-Object Id -Unique
            
            # Log all found processes
            foreach ($process in $allCandidateProcesses) {
                # Double-check if process is still running
                if ($process.HasExited) { continue }

                $processInfo = [PSCustomObject]@{
                    Name = $process.ProcessName
                    Id = $process.Id
                    Title = $process.MainWindowTitle # May be empty
                    Process = $process
                }
                $userProcessesToClose += $processInfo
                Write-LogMessage "Found candidate: $($processInfo.Name) (PID: $($processInfo.Id), Title: '$($processInfo.Title)')" -Level Debug
            }
        } catch {
            Write-LogMessage "Error occurred while scanning for user processes: $($_)" -Level Error
            $closeSuccess = $false # Mark as failure if scan fails
            # Continue if possible, but report error
        }

        if (-not $closeSuccess) {
            Write-LogMessage "Aborting application closing due to previous error during scan." -Level Error
            return $false
        }

        if ($userProcessesToClose.Count -eq 0) {
            Write-LogMessage "No user applications found requiring closure." -Level Success
            return $true # Nothing to close
        }
        # --- End Get User Applications ---


        # --- Display and Countdown ---
        Write-LogMessage "`nApplications identified for closure:" -Level Info
        $i = 1
        foreach ($app in $userProcessesToClose) {
            $displayName = "$($app.Name) (PID: $($app.Id))"
            $displayTitle = if ([string]::IsNullOrWhiteSpace($app.Title)) { "" } else { " - $($app.Title)" }
            Write-Host ("{0}. {1}{2}" -f $i, $displayName, $displayTitle) -ForegroundColor 'White'
            $i++
        }

        $effectiveForce = $ForceParam -or $script:config.AlwaysForce
        if (-not $effectiveForce) {
            $actionColor = if ($script:config.colors.action) { $script:config.colors.action } else { 'Magenta' }
            $warningColor = if ($script:config.colors.warning) { $script:config.colors.warning } else { 'Yellow' }
            Write-Host "`nPreparing to close applications and perform '$Action'" -ForegroundColor $actionColor
            Write-Host "Operation will begin in $($TimeoutMSParam / 1000) seconds. Press Ctrl+C to cancel." -ForegroundColor $warningColor

            try {
                for ($i = ($TimeoutMSParam / 1000); $i -gt 0; $i--) {
                    Write-Host ("Starting in {0}..." -f $i) -ForegroundColor $warningColor -NoNewline
                    Start-Sleep -Seconds 1
                    # Overwrite line using carriage return
                    Write-Host "`r                                                `r" -NoNewline
                }
                 Write-LogMessage "Countdown complete. Starting operation now..." -Level Info
            } catch [System.Management.Automation.PipelineStoppedException] {
                 Write-LogMessage "Operation cancelled by user (Ctrl+C)." -Level Warning
                 return $false # User cancelled
            } catch {
                Write-LogMessage "Error during countdown: ${_}. Proceeding with closure." -Level Warning
            }
        } else {
            Write-LogMessage "Force mode active. Proceeding immediately with application closure." -Level Info
        }
        # --- End Display and Countdown ---


        # --- Close Processes ---
        Write-LogMessage "`nClosing applications..." -Level Action
        $effectiveNoGraceful = $script:config.NoGraceful
        $actualTimeout = [Math]::Max(1000, $TimeoutMSParam) # Ensure at least 1 sec timeout

        # Pre-assign colors to local variables
        $warningColor = if ($script:config.colors.warning) { $script:config.colors.warning } else { 'Yellow' }
        $successColor = if ($script:config.colors.success) { $script:config.colors.success } else { 'Green' }
        $errorColor = if ($script:config.colors.error) { $script:config.colors.error } else { 'Red' }

        # First attempt - standard close method
        foreach ($app in $userProcessesToClose) {
            $appName = "$($app.Name) (PID: $($app.Id))"
            Write-Host "Closing $appName" -NoNewline -ForegroundColor 'White'
            if (-not [string]::IsNullOrWhiteSpace($app.Title)) { Write-Host " - $($app.Title)" -NoNewline -ForegroundColor 'DarkGray'}
            Write-Host "... " -NoNewline

            $processClosed = $false
            try {
                # Check if process still exists before attempting close
                $currentProcess = Get-Process -Id $app.Id -ErrorAction SilentlyContinue
                if (-not $currentProcess -or $currentProcess.HasExited) {
                    Write-Host "Already exited." -ForegroundColor 'Gray'
                    $processClosed = $true
                    continue # Skip to next app
                }

                # Try to give focus to the window if it has one
                if ($currentProcess.MainWindowHandle -ne [System.IntPtr]::Zero) {
                    try {
                        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class WindowManagement {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@ -ErrorAction SilentlyContinue

                        [WindowManagement]::SetForegroundWindow($currentProcess.MainWindowHandle) | Out-Null
                        Start-Sleep -Milliseconds 100
                    } catch {
                        # Ignore focus errors - just trying to improve graceful close
                        Write-LogMessage "Failed to set focus to window: $_" -Level Debug
                    }
                }

                # Determine if force kill is needed immediately
                $forceKill = $effectiveForce -or $effectiveNoGraceful

                if (-not $forceKill) {
                    # Attempt graceful close
                    Write-Host "Graceful... " -NoNewline -ForegroundColor 'Gray'
                    $closedGracefully = $app.Process.CloseMainWindow()

                    if ($closedGracefully) {
                        # Wait for process to exit after graceful signal
                        $exited = $app.Process.WaitForExit($actualTimeout)
                        if ($exited) {
                            Write-Host "OK." -ForegroundColor $successColor
                            $processClosed = $true
                        } else {
                            # Timed out after graceful attempt, now force
                             Write-Host "Timeout, forcing... " -NoNewline -ForegroundColor $warningColor
                             $forceKill = $true
                        }
                    } else {
                        # Graceful close failed immediately (e.g., no window reaction), force needed
                        Write-Host "No response, forcing... " -NoNewline -ForegroundColor $warningColor
                        $forceKill = $true
                    }
                }

                # Force kill if needed (initial force, timeout, or graceful fail)
                if ($forceKill -and -not $processClosed) {
                     Write-Host "Forcing... " -NoNewline -ForegroundColor $warningColor
                     
                     # First try normal kill
                     $app.Process.Kill()
                     
                     # Short wait after kill
                     Start-Sleep -Milliseconds 200
                     
                     # Verify exit after kill
                     $processCheck = Get-Process -Id $app.Id -ErrorAction SilentlyContinue
                     if ($null -eq $processCheck -or $processCheck.HasExited) {
                         Write-Host "OK." -ForegroundColor $successColor
                         $processClosed = $true
                     } else {
                         # Try more aggressive kill with taskkill
                         Write-Host "Taskkill... " -NoNewline -ForegroundColor $warningColor
                         $taskKillResult = Start-Process -FilePath "taskkill.exe" -ArgumentList "/F /PID $($app.Id)" -Wait -NoNewWindow -PassThru
                         
                         # Verify one more time
                         Start-Sleep -Milliseconds 200
                         $finalCheck = Get-Process -Id $app.Id -ErrorAction SilentlyContinue
                         if ($null -eq $finalCheck -or $finalCheck.HasExited) {
                             Write-Host "OK." -ForegroundColor $successColor
                             $processClosed = $true
                         } else {
                             Write-Host "FAILED." -ForegroundColor $errorColor
                             $processClosed = $false
                         }
                     }
                }

            } catch {
                 Write-Host "Error: $($_.Exception.Message)" -ForegroundColor $errorColor
                 $processClosed = $false # Mark as failed on error
            } finally {
                if (-not $processClosed) {
                    $failedProcesses += $appName
                    Write-LogMessage "Failed to close: $appName" -Level Error
                }
            }
        }

        # Second attempt for failed processes - more aggressive approach
        if ($failedProcesses.Count -gt 0) {
            Write-LogMessage "Attempting more aggressive termination for stubborn processes..." -Level Warning
            
            # Clear the failed processes array for second attempt tracking
            $secondAttemptFailures = @()
            
            foreach ($processName in $failedProcesses) {
                $match = $processName -match '^(.*?)\s\(PID:\s(\d+)\)'
                if ($match) {
                    $name = $Matches[1]
                    $id = [int]$Matches[2]
                    
                    Write-LogMessage "Second attempt to terminate $name (PID: $id)..." -Level Warning
                    
                    try {
                        # Try direct Windows API termination for more forceful termination
                        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

public class ProcessTermination
{
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr OpenProcess(uint processAccess, bool bInheritHandle, int processId);
    
    [DllImport("kernel32.dll", SetLastError=true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool TerminateProcess(IntPtr hProcess, uint uExitCode);
    
    [DllImport("kernel32.dll", SetLastError=true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool CloseHandle(IntPtr hObject);
    
    public static bool ForceTerminate(int processId)
    {
        IntPtr hProcess = OpenProcess(0x0001, false, processId); // PROCESS_TERMINATE
        if (hProcess == IntPtr.Zero)
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }
        
        try
        {
            if (!TerminateProcess(hProcess, 1))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            return true;
        }
        finally
        {
            CloseHandle(hProcess);
        }
    }
}
"@ -ErrorAction Stop

                        $terminated = [ProcessTermination]::ForceTerminate($id)
                        if ($terminated) {
                            Write-LogMessage "Successfully terminated $name with Windows API" -Level Success
                        } else {
                            # Fallback to taskkill /F /T (kill tree)
                            $result = Start-Process -FilePath "taskkill.exe" -ArgumentList "/F /T /PID $id" -Wait -NoNewWindow -PassThru
                            
                            if ($result.ExitCode -eq 0) {
                                Write-LogMessage "Successfully terminated $name with taskkill /F /T" -Level Success
                            } else {
                                Write-LogMessage "Taskkill failed with exit code $($result.ExitCode)" -Level Error
                                $secondAttemptFailures += $processName
                            }
                        }
                    } catch {
                        Write-LogMessage "Error using process termination API: $_" -Level Error
                        # Final fallback to the most aggressive taskkill method
                        try {
                            # Try without the complex filter that's causing linter issues
                            $result = Start-Process -FilePath "taskkill.exe" -ArgumentList "/F /T /PID $id" -Wait -NoNewWindow -PassThru
                            if ($result.ExitCode -eq 0) {
                                Write-LogMessage "Successfully terminated $name with emergency fallback method" -Level Success
                            } else {
                                $secondAttemptFailures += $processName
                            }
                        } catch {
                            Write-LogMessage "All termination methods failed for: $name" -Level Error
                            $secondAttemptFailures += $processName
                        }
                    }
                } else {
                    Write-LogMessage "Could not parse process name from: $processName" -Level Error
                    $secondAttemptFailures += $processName
                }
            }
            
            # Update the failed processes list with only those that failed both attempts
            $failedProcesses = $secondAttemptFailures
        }
        # --- End Close Processes ---

        # --- Report Results ---
        $elapsedTime = (Get-Date) - $startTime
        Write-LogMessage "Process closure operation completed in $($elapsedTime.TotalSeconds) seconds" -Level Info
        
        if ($failedProcesses.Count -gt 0) {
            $failedList = $failedProcesses -join ', '
            Write-LogMessage "Failed to close the following application(s): $failedList" -Level Error
            
            # Return failure unless Force was used
            $finalResult = $effectiveForce
            $resultMessage = if ($finalResult) { "Success (force override)" } else { "Failure" }
            Write-LogMessage "Final result: $resultMessage" -Level Info
            return $finalResult
        } else {
            Write-LogMessage "Successfully closed all identified applications." -Level Success
            return $true
        }

    } catch {
        Write-LogMessage "Unexpected error during Close-UserProcesses: $_" -Level Error
        if ($ErrorState) {
            try { $ErrorState.Value = $_ } catch {} # Report error back if ref provided
        }
        return $false # Return failure on unexpected error
    } finally {
        Write-LogMessage "============ CLOSE PROCESSES SESSION END ============" -Level Info
    }
}


#region Power Action Functions (Refactored)

# Private helper to execute power actions with retries
function Invoke-PowerAction {
    param(
        [ValidateSet('Sleep', 'Restart', 'Shutdown')]
        [string]$ActionType
    )

    Write-LogMessage "Initiating $ActionType sequence..." -Level Action
    Start-Sleep -Seconds 1 # Brief pause before action

    $success = $false
    $lastError = $null

    # Check for administrator privileges
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-LogMessage "Warning: Not running with administrator privileges. Some power actions may fail." -Level Warning
    }

    # Define methods for each action type
    $actionMethods = @{
        Sleep = @(
            @{ Name="API Call"; Command={ 
                Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class SleepManager {
    [DllImport("PowrProf.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
    public static extern bool SetSuspendState(bool hibernate, bool forceCritical, bool disableWakeEvent);
}
"@ -ErrorAction SilentlyContinue
                # Get last Win32 error code if available
                $errorCode = 0
                $result = [SleepManager]::SetSuspendState($false, $true, $false)
                if (-not $result) {
                    $errorCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
                    Write-LogMessage "API call returned error code: $errorCode" -Level Warning
                    return $false
                }
                return $true
            }; PreCheck=$false; PostDelay=3 },
            @{ Name="PowerProf DLL"; Command={ 
                $proc = Start-Process -FilePath "rundll32.exe" -ArgumentList "powrprof.dll,SetSuspendState 0,1,0" -PassThru -Wait -NoNewWindow
                return ($proc.ExitCode -eq 0)
            }; PreCheck=$true; PostDelay=3 },
            @{ Name="Hibernate Fallback"; Command={ 
                $proc = Start-Process -FilePath "shutdown.exe" -ArgumentList "/h /f /t 0" -PassThru -Wait -NoNewWindow
                return ($proc.ExitCode -eq 0)
            }; PreCheck=$true; PostDelay=3 } # Requires hibernate enabled
        )
        Restart = @(
            @{ Name="shutdown.exe /r"; Command={ 
                $proc = Start-Process -FilePath "shutdown.exe" -ArgumentList "/r /f /t 0" -PassThru -Wait -NoNewWindow
                return ($proc.ExitCode -eq 0)
            }; PreCheck=$true; PostDelay=2 },
            @{ Name="PowerShell Restart-Computer"; Command={ 
                Restart-Computer -Force -ErrorAction Stop
                # If we get here, the restart didn't happen immediately
                Start-Sleep -Seconds 2
                return $false
            }; PreCheck=$false; PostDelay=3 },
            @{ Name="Start-Process shutdown /r"; Command={ 
                $proc = Start-Process -FilePath "shutdown.exe" -ArgumentList "/r /f /t 0" -PassThru -Wait -NoNewWindow
                return ($proc.ExitCode -eq 0)
            }; PreCheck=$true; PostDelay=0 }
        )
        Shutdown = @(
            @{ Name="shutdown.exe /s"; Command={ 
                $proc = Start-Process -FilePath "shutdown.exe" -ArgumentList "/s /f /t 0" -PassThru -Wait -NoNewWindow
                return ($proc.ExitCode -eq 0)
            }; PreCheck=$true; PostDelay=2 },
            @{ Name="PowerShell Stop-Computer"; Command={ 
                Stop-Computer -Force -ErrorAction Stop
                # If we get here, the shutdown didn't happen immediately
                Start-Sleep -Seconds 2
                return $false
            }; PreCheck=$false; PostDelay=3 },
            @{ Name="Start-Process shutdown /s"; Command={ 
                $proc = Start-Process -FilePath "shutdown.exe" -ArgumentList "/s /f /t 0" -PassThru -Wait -NoNewWindow
                return ($proc.ExitCode -eq 0)
            }; PreCheck=$true; PostDelay=0 }
        )
    }

    # Log the current power capabilities 
    try {
        $powerStatus = Get-WmiObject -Class Win32_PowerCapabilities -Namespace 'root\cimv2\power' -ErrorAction SilentlyContinue
        if ($powerStatus) {
            Write-LogMessage "System power capabilities: S1=$($powerStatus.SystemS1), S2=$($powerStatus.SystemS2), S3=$($powerStatus.SystemS3), S4=$($powerStatus.SystemS4)" -Level Debug
        }
    } catch {
        Write-LogMessage "Unable to retrieve power capabilities: $_" -Level Debug
    }

    # Pre-action steps for Sleep
    if ($ActionType -eq 'Sleep') {
        try {
            # Make sure display won't immediately wake up the system
            Write-LogMessage "Disabling wake timers..." -Level Debug
            powercfg /setacvalueindex SCHEME_CURRENT 238c9fa8-0aad-41ed-83f4-97be242c8f20 bd3b718a-0680-4d9d-8ab2-e1d2b4ac806d 0
            powercfg /setdcvalueindex SCHEME_CURRENT 238c9fa8-0aad-41ed-83f4-97be242c8f20 bd3b718a-0680-4d9d-8ab2-e1d2b4ac806d 0
            
            # Apply the power settings
            powercfg /setactive SCHEME_CURRENT
            
            # Enable S3 sleep if available
            $s3State = powercfg /a | Select-String -Pattern "The following sleep states are available" -Context 0, 5
            if ($s3State -and $s3State.Context.PostContext -match "S3") {
                Write-LogMessage "S3 sleep state is available" -Level Debug
            } else {
                Write-LogMessage "Warning: S3 sleep state might not be available" -Level Warning
            }
        } catch { 
            Write-LogMessage "Warning: Failed to prepare system for sleep: $_" -Level Warning 
        }
    }

    # Log detailed system info
    Write-LogMessage "System information:" -Level Info
    Write-LogMessage "OS: $([System.Environment]::OSVersion.VersionString)" -Level Info
    Write-LogMessage "PowerShell: $($PSVersionTable.PSVersion)" -Level Info
    Write-LogMessage "Admin: $isAdmin" -Level Info

    # Try methods sequentially
    foreach ($method in $actionMethods[$ActionType]) {
        Write-LogMessage "Attempting $ActionType via: $($method.Name)" -Level Info
        try {
            # Pre-check if needed (e.g., ensure command exists)
            if ($method.PreCheck) {
                # For commands like rundll32, shutdown, check existence in System32
                $cmd = ($method.Command.ToString() -split ' ')[0]
                $cmdPath = Join-Path $env:SystemRoot "System32" $cmd
                 if (-not (Test-PathExists -Path $cmdPath -PathType Leaf)) {
                     Write-LogMessage "Command '$cmd' not found at '$cmdPath'. Skipping method." -Level Warning
                     continue
                 }
            }

            # Log that we're trying this method
            Write-LogMessage "Executing command: $($method.Command)" -Level Info

            # Execute the command/scriptblock with a timeout
            $scriptBlock = $method.Command
            $job = Start-Job -ScriptBlock { 
                param($sb)
                & $sb
            } -ArgumentList $scriptBlock

            # Wait for command completion with timeout
            $jobCompleted = Wait-Job -Job $job -Timeout 10 -ErrorAction SilentlyContinue
            
            if ($null -eq $jobCompleted) {
                Write-LogMessage "Command timed out after 10 seconds" -Level Warning
                Stop-Job -Job $job -ErrorAction SilentlyContinue
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                $result = $false
            } else {
                $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            }

            # If command returns $true or command doesn't throw and doesn't immediately exit script
            if ($result -eq $true) {
                $success = $true
                Write-LogMessage "$($method.Name) command returned success." -Level Success
                break # Exit loop on success
            }

            # Wait briefly to see if the action took effect
            Write-LogMessage "Waiting to see if $ActionType occurs..." -Level Info
            Start-Sleep -Seconds $method.PostDelay

            # If we're still here, the action didn't work
            Write-LogMessage "$($method.Name) didn't appear to succeed. Trying next method." -Level Warning

        } catch {
            $lastError = $_
            Write-LogMessage "Method '$($method.Name)' failed: $lastError" -Level Warning
            # Wait briefly before next attempt
            Start-Sleep -Seconds 1
        }
    }

    # If all methods failed, try some recovery steps
    if (-not $success) {
        Write-LogMessage "All standard methods failed to initiate $ActionType. Trying recovery steps..." -Level Warning
        
        try {
            # For Sleep failures, try direct WMI call
            if ($ActionType -eq 'Sleep') {
                Write-LogMessage "Attempting WMI method..." -Level Info
                # More direct approach to sleep using Win32 API
                Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class PowerState
{
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern uint SetThreadExecutionState(uint esFlags);

    [DllImport("Powrprof.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
    public static extern bool SetSuspendState(bool hibernate, bool forceCritical, bool disableWakeEvent);

    public static void ForceSleep()
    {
        SetThreadExecutionState(0x00000001); // ES_CONTINUOUS
        SetSuspendState(false, true, false);
    }
}
"@
                [PowerState]::ForceSleep()
                Start-Sleep -Seconds 3 # Wait to see if it worked
                $success = $true
            }
            
            # For Restart/Shutdown, try psshutdown.exe if available
            elseif (($ActionType -eq 'Restart' -or $ActionType -eq 'Shutdown')) {
                # Try with even more force using shutdown.exe with different parameters
                $shutdownParams = if ($ActionType -eq 'Restart') { "/r /f /t 0" } else { "/s /f /t 0" }
                Write-LogMessage "Attempting shutdown with additional force flags..." -Level Info
                
                # First kill the explorer process to ensure it doesn't block
                Get-Process -Name explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 1
                
                # Execute shutdown directly without Start-Process
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = "$env:SystemRoot\System32\shutdown.exe"
                $psi.Arguments = $shutdownParams
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $true
                
                $process = [System.Diagnostics.Process]::Start($psi)
                $success = $true
            }
        } catch {
            Write-LogMessage "Recovery attempt failed: $_" -Level Error
        }
    }

    if (-not $success) {
        Write-LogMessage "All methods failed to initiate $ActionType. Last error: $lastError" -Level Error
        Write-LogMessage "Please try again with -Force or contact system administrator" -Level Error
    }

    return $success
}

#endregion

#region MAC Address Formatting Function

function Format-MACAddress {
    param (
        [string]$MacAddressInput
    )

    if ([string]::IsNullOrWhiteSpace($MacAddressInput)) {
        throw "MAC address cannot be empty."
    }

    # Remove common separators and convert to lowercase hex characters only
    $cleanedMac = ($MacAddressInput -replace '[:-.\s]', '').ToLowerInvariant()

    # Validate format (12 hexadecimal characters)
    if ($cleanedMac -notmatch '^[a-f0-9]{12}$') {
        throw "Invalid MAC address format: '$MacAddressInput'. Must contain 12 hexadecimal characters (0-9, a-f)."
    }

    # Format with colons
    # Example: "0123456789ab" -> "01", "23", "45", "67", "89", "ab" -> "01:23:45:67:89:ab"
    $formattedMac = ($cleanedMac -split '([a-f0-9]{2})' | Where-Object { $_ }) -join ':'

    return $formattedMac
}

#endregion

#region Main Execution

# --- Parameter Set Handling ---
if ($PSCmdlet.ParameterSetName -eq 'FormatMAC') {
    try {
        $formatted = Format-MACAddress -MacAddressInput $FormatMAC
        Write-Host $formatted # Output directly for pcpow.bat
        exit 0
    } catch {
        Write-Error "MAC Formatting Error: $($_.Exception.Message)"
        exit 1
    }
}

# --- Proceed with PowerAction Parameter Set ---
$script:startTime = Get-Date
$script:sessionID = [Guid]::NewGuid().ToString("N")

# Check for administrator privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and -not $Force) {
    # Only warn if not forced - we'll try anyway if forced
    Write-LogMessage "Warning: Not running with administrator privileges. Power actions may require elevation." -Level Warning
    
    # For Sleep action specifically, check if it's likely to work without admin rights
    if ($Action -eq "Sleep") {
        $sleepStates = powercfg /a | Out-String
        if ($sleepStates -notmatch "Standby \(S3\) is available") {
            Write-LogMessage "S3 sleep state may not be available. This might cause problems without admin rights." -Level Warning
        }
    }
}

# Verify that we're not running from a network drive or restricted location
$scriptPath = $PSCommandPath
if ($scriptPath -and $scriptPath -match "^\\\\") {
    Write-LogMessage "Warning: Running script from a network location ($scriptPath) which may affect power management capabilities." -Level Warning
}

# Basic system validity checks
if (-not (Get-Command "shutdown.exe" -ErrorAction SilentlyContinue)) {
    Write-LogMessage "Critical error: shutdown.exe not found in PATH. System environment may be corrupt." -Level Error
    exit 1
}

# Initialize logging
Write-LogMessage "PCPow v$($script:config.version) Initializing (Session: $script:sessionID)..." -Level Info # Use version from config

# Start logging session info
Write-LogMessage "==================== SESSION START ====================" -Level Info
Write-LogMessage "Session ID: $script:sessionID" -Level Info
Write-LogMessage "Date/Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level Info
Write-LogMessage "User: $($env:USERNAME) | Machine: $($env:COMPUTERNAME)" -Level Info
Write-LogMessage "Action: $Action | Force: $($Force)" -Level Info
Write-LogMessage "PSVersion: $($PSVersionTable.PSVersion)" -Level Info
Write-LogMessage "WorkingDir: $(Get-Location)" -Level Info
Write-LogMessage "Admin Rights: $isAdmin" -Level Info
Write-LogMessage "Script Path: $scriptPath" -Level Info

# Register a script end block to ensure we always log completion
Register-EngineEvent -SourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) -Action {
    $exitTime = Get-Date
    $duration = $exitTime - $script:startTime
    Write-LogMessage "Session $script:sessionID completed in $($duration.TotalSeconds) seconds" -Level Info
    Write-LogMessage "==================== SESSION END ====================" -Level Info
} -ErrorAction SilentlyContinue

# Verify running in a valid environment
try {
    if (-not $env:SystemRoot -or -not (Test-PathExists -Path $env:SystemRoot -PathType Container)) {
         throw "Windows SystemRoot environment variable not found or invalid."
    }
    $currentDir = Get-Location -ErrorAction Stop
    if (-not (Test-PathExists -Path $currentDir.Path -PathType Container)) {
        throw "Current working directory '$($currentDir.Path)' is invalid."
    }
     Write-LogMessage "Environment checks passed (SystemRoot: $env:SystemRoot, CWD: $($currentDir.Path))" -Level Debug
} catch {
     Write-LogMessage "Critical environment error: $_. Aborting." -Level Error
     exit 1
}


# Load configuration (script:config will be populated)
$null = Get-ConfigurationSettings # Load/validate config, suppress output here

# Log action
$forceModeString = if ($Force -or $script:config.AlwaysForce) { "(Force Mode Active)" } else { "" }
Write-LogMessage "Requested Action: $Action $forceModeString" -Level Action

# Check for potential blocking applications
try {
    $potentialBlockers = @(
        "MicrosoftEdgeUpdate",
        "OneDrive",
        "SamsungMagician", 
        "NVDisplay.Container", 
        "AdobeUpdateService",
        "BackgroundTaskHost",
        "WindowsUpdateAgent",
        "CCleaner",
        "iTunesHelper"
    )
    
    $runningBlockers = @()
    foreach ($blocker in $potentialBlockers) {
        $process = Get-Process -Name $blocker -ErrorAction SilentlyContinue
        if ($process) {
            $runningBlockers += "$blocker (PID: $($process.Id))"
        }
    }
    
    if ($runningBlockers.Count -gt 0) {
        Write-LogMessage "Potential blocking processes detected: $($runningBlockers -join ', ')" -Level Warning
        Write-LogMessage "These processes may interfere with $Action operation. Consider adding them to excludedProcesses in config." -Level Warning
    }
} catch {
    Write-LogMessage "Error checking for potential blocking processes: $_" -Level Warning
}

# Log current running processes for diagnostics
try {
    $importantProcesses = Get-Process | Where-Object { $_.MainWindowHandle -ne [System.IntPtr]::Zero } | 
                         Select-Object Name, Id, MainWindowTitle | 
                         Format-Table -AutoSize | Out-String -Width 120
    Write-LogMessage "Currently running processes with windows:`n$importantProcesses" -Level Debug
} catch {
    Write-LogMessage "Unable to log current processes: $_" -Level Warning
}

$global:ErrorState = $null # Variable to capture errors from Close-UserProcesses

# Create consolidated excluded processes list with explicit case-variation checks for extra safety
$excludedProcessesList = $script:config.excludedProcesses + @(
    # Add the current PowerShell host process
    (Get-Process -Id $PID).ProcessName,
    # Add its parent process if it's a terminal
    (Get-Process -Id (Get-CimInstance Win32_Process -Filter "ProcessId = $PID").ParentProcessId -ErrorAction SilentlyContinue).ProcessName
) | Where-Object { $null -ne $_ } | Select-Object -Unique

# Close applications
$closeSuccess = Close-UserProcesses -ForceParam:($Force) `
                                    -TimeoutMSParam $script:config.TimeoutMS `
                                    -ExcludedProcessesParam $excludedProcessesList `
                                    -ErrorState ([ref]$global:ErrorState)

if (-not $closeSuccess) {
    if ($global:ErrorState) {
        Write-LogMessage "Application closing failed due to error: $($global:ErrorState.Exception.Message)" -Level Error
    } else {
        Write-LogMessage "Application closing failed. Check logs above." -Level Error
    }
    # Exit if not forcing
    if (-not ($Force -or $script:config.AlwaysForce)) {
        Write-LogMessage "Aborting $Action operation due to application close failure. Use -Force to override." -Level Error
        exit 1
    } else {
        Write-LogMessage "Force mode active. Proceeding with $Action despite application close failures." -Level Warning
    }
}

# Skip power action if requested
if ($SkipAction) {
    Write-LogMessage "Skipping $Action operation as requested (-SkipAction)." -Level Warning
    Write-LogMessage "PCPow finished in $( (Get-Date) - $startTime ).TotalSeconds seconds." -Level Info
    exit 0
}

# Perform final preparation before power action
try {
    # Make sure nothing is holding files open in user profile
    Write-LogMessage "Clearing standard input/output buffers..." -Level Debug
    [System.Console]::Clear()
    [System.Console]::Out.Flush()
    [System.Console]::Error.Flush()
    
    # Clear any pending keyboard or mouse input
    [System.Console]::In.ReadToEnd() | Out-Null
    
    # Force garbage collection to release any COM objects or file handles
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    
    # Brief pause to let everything settle
    Start-Sleep -Seconds 1
    
    Write-LogMessage "Final preparations complete. Executing $Action..." -Level Action
} catch {
    Write-LogMessage "Warning during final preparations: $_" -Level Warning
    # Continue anyway - this is just cleanup
}

# Perform requested power action
$actionSuccess = Invoke-PowerAction -ActionType $Action

if (-not $actionSuccess) {
    Write-LogMessage "Failed to perform $Action action. Check logs for details." -Level Error
     Write-LogMessage "PCPow finished with errors in $( (Get-Date) - $startTime ).TotalSeconds seconds." -Level Error
    exit 1
}

# Note: Successful power actions (Sleep, Restart, Shutdown) will likely terminate the script before this point.
# This exit is a fallback.
Write-LogMessage "$Action initiated successfully." -Level Success
Write-LogMessage "PCPow finished in $( (Get-Date) - $startTime ).TotalSeconds seconds." -Level Info
exit 0

#endregion 