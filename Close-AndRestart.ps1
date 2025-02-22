# Close applications and restart PC
param (
    [switch]$Force = $false
)

function Show-ConfirmationPrompt {
    $title = "Close All Applications"
    $message = "WARNING: This script will attempt to close all running programs and restart the PC.`nPlease save any important work before continuing.`nDo you want to proceed?"
    
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Close all applications and restart"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Cancel the operation"
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
        [System.Diagnostics.Process[]]$Processes
    )
    
    foreach ($proc in $Processes) {
        try {
            Write-Host "Attempting to close $($proc.Name)..."
            if ($proc.CloseMainWindow()) {
                # Wait up to 5 seconds for the process to close gracefully
                if (!$proc.WaitForExit(5000)) {
                    if ($Force) {
                        Write-Host "Force closing $($proc.Name)..."
                        $proc | Stop-Process -Force
                    }
                }
            } else {
                if ($Force) {
                    Write-Host "Force closing $($proc.Name)..."
                    $proc | Stop-Process -Force
                }
            }
        } catch {
            Write-Warning "Failed to close $($proc.Name): $_"
        }
    }
}

# Main script execution
if (-not $Force) {
    if (-not (Show-ConfirmationPrompt)) {
        Write-Host "Operation cancelled by user."
        exit
    }
}

Write-Host "Identifying running applications..."
$userApps = Get-UserApps

if ($userApps.Count -eq 0) {
    Write-Host "No user applications found to close."
} else {
    Write-Host "Found $($userApps.Count) applications to close."
    Close-Apps -Processes $userApps
}

Write-Host "Waiting for processes to finish closing..."
Start-Sleep -Seconds 2

Write-Host "Restarting PC..."
if ($Force) {
    Restart-Computer -Force
} else {
    Restart-Computer
} 