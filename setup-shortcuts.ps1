# Setup script to install shortcuts and aliases
$ErrorActionPreference = 'Stop'

# Get the directory where the scripts are located
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = [System.IO.Path]::GetFullPath($scriptPath)

# Create module directory in PowerShell modules path
$moduleName = "pcpow-common"
$moduleVersion = "1.0.0"
# Use PowerShell's built-in module path
$userModulePath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "WindowsPowerShell\Modules\pcpow-common\$moduleVersion"

Write-Host "Creating module directory: $userModulePath"
if (-not (Test-Path $userModulePath)) {
    New-Item -ItemType Directory -Path $userModulePath -Force | Out-Null
}

# Create module manifest
$manifestPath = Join-Path $userModulePath "$moduleName.psd1"
Write-Host "Creating module manifest..."
New-ModuleManifest -Path $manifestPath `
    -RootModule "$moduleName.psm1" `
    -ModuleVersion $moduleVersion `
    -Author "PCPow Team" `
    -Description "PCPow power management module" `
    -PowerShellVersion "5.1"

# Create shortcuts directory in the Windows directory if it doesn't exist
$shortcutsDir = Join-Path $env:USERPROFILE "AppData\Local\Microsoft\WindowsApps"
if (-not (Test-Path $shortcutsDir)) {
    New-Item -ItemType Directory -Path $shortcutsDir -Force | Out-Null
}

# Copy module files
Write-Host "Installing PowerShell module..."
Write-Host "Copying $scriptPath\pcpow-common.psm1 to $userModulePath\$moduleName.psm1"
Copy-Item "$scriptPath\pcpow-common.psm1" "$userModulePath\$moduleName.psm1" -Force

# Copy config file to both locations for redundancy
Write-Host "Copying configuration file..."
Copy-Item "$scriptPath\pcpow.config.json" "$userModulePath\" -Force
Copy-Item "$scriptPath\pcpow.config.json" "$shortcutsDir\" -Force

# Verify module files
if (-not (Test-Path "$userModulePath\$moduleName.psm1")) {
    throw "Failed to copy module file to $userModulePath\$moduleName.psm1"
}
if (-not (Test-Path "$userModulePath\pcpow.config.json")) {
    throw "Failed to copy config file to $userModulePath\pcpow.config.json"
}
if (-not (Test-Path "$shortcutsDir\pcpow.config.json")) {
    throw "Failed to copy config file to $shortcutsDir\pcpow.config.json"
}
if (-not (Test-Path $manifestPath)) {
    throw "Failed to create module manifest at $manifestPath"
}

# Copy script files
Write-Host "Copying script files..."
Write-Host "Copying pcpow.bat to $shortcutsDir"
Copy-Item "$scriptPath\pcpow.bat" $shortcutsDir -Force
if (-not (Test-Path "$shortcutsDir\pcpow.bat")) {
    throw "Failed to copy pcpow.bat to $shortcutsDir"
}

Write-Host "Copying PowerShell scripts..."
foreach ($script in @("Close-AndSleep.ps1", "Close-AndRestart.ps1", "Close-AndShutdown.ps1")) {
    Write-Host "  Copying $script"
    Copy-Item "$scriptPath\$script" $shortcutsDir -Force
    if (-not (Test-Path "$shortcutsDir\$script")) {
        throw "Failed to copy $script to $shortcutsDir"
    }
}

# Verify the copied batch file is correct
$batchContent = Get-Content "$shortcutsDir\pcpow.bat" -Raw
if (-not $batchContent.Contains("Start-PCSleep")) {
    Write-Warning "pcpow.bat may not be up to date. Please verify the help text is current."
}

# Create PowerShell profile directory if it doesn't exist
$profileDir = Split-Path -Parent $PROFILE
if (-not (Test-Path $profileDir)) {
    New-Item -ItemType Directory -Path $profileDir -Force | Out-Null
}

# Create PowerShell profile if it doesn't exist
if (-not (Test-Path $PROFILE)) {
    New-Item -ItemType File -Path $PROFILE -Force | Out-Null
}

# Remove old aliases if they exist
Write-Host "Cleaning up PowerShell profile..."
if (Test-Path $PROFILE) {
    $profileContent = Get-Content $PROFILE -Raw
    # Remove existing PCPow blocks using regex
    $profileContent = $profileContent -replace '(?s)# PCPow Start.*?# PCPow End', ''
    Set-Content $PROFILE $profileContent
}

# Modify the $moduleConfig to include markers
$moduleConfig = @"
# PCPow Start
# PCPow Configuration
`$ErrorActionPreference = 'Stop'

# Add module path if needed
`$modulePath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Modules'
if (-not (`$env:PSModulePath -split ';' -contains `$modulePath)) {
    `$env:PSModulePath = `$env:PSModulePath + ';' + `$modulePath
}

# Import PCPow module
Import-Module pcpow-common -MinimumVersion $moduleVersion -Force -ErrorAction Stop

# PCPow - Power Management Functions
function global:Start-PCSleep {
    param(
        [switch]`$Force
    )
    try {
        Import-Module pcpow-common -ErrorAction Stop
        `$config = Get-Content "$shortcutsDir\pcpow.config.json" | ConvertFrom-Json
        `$useForce = `$Force -or `$config.AlwaysForce
        & "$shortcutsDir\Close-AndSleep.ps1" -Force:`$useForce
    } catch {
        Write-Warning "Error in Start-PCSleep: `$_"
        & "$shortcutsDir\Close-AndSleep.ps1" -Force:`$Force
    }
}

function global:Restart-PC {
    param(
        [switch]`$Force
    )
    try {
        Import-Module pcpow-common -ErrorAction Stop
        `$config = Get-Content "$shortcutsDir\pcpow.config.json" | ConvertFrom-Json
        `$useForce = `$Force -or `$config.AlwaysForce
        & "$shortcutsDir\Close-AndRestart.ps1" -Force:`$useForce
    } catch {
        Write-Warning "Error in Restart-PC: `$_"
        & "$shortcutsDir\Close-AndRestart.ps1" -Force:`$Force
    }
}

function global:Stop-PC {
    param(
        [switch]`$Force
    )
    try {
        Import-Module pcpow-common -ErrorAction Stop
        `$config = Get-Content "$shortcutsDir\pcpow.config.json" | ConvertFrom-Json
        `$useForce = `$Force -or `$config.AlwaysForce
        & "$shortcutsDir\Close-AndShutdown.ps1" -Force:`$useForce
    } catch {
        Write-Warning "Error in Stop-PC: `$_"
        & "$shortcutsDir\Close-AndShutdown.ps1" -Force:`$Force
    }
}

# Create aliases
Set-Alias -Name pows -Value Start-PCSleep -Scope Global
Set-Alias -Name powr -Value Restart-PC -Scope Global
Set-Alias -Name powd -Value Stop-PC -Scope Global

Write-Host "PCPow commands loaded successfully" -ForegroundColor Green
# PCPow End
"@

Add-Content -Path $PROFILE -Value $moduleConfig

# Test the commands
Write-Host "`nTesting command availability..."
$testScript = @"
`$ErrorActionPreference = 'Stop'
try {
    Import-Module pcpow-common -Force
    Write-Host "Available commands:"
    Write-Host "  pcpow sleep/restart/shutdown"
    Write-Host "  pows, powr, powd"
    Write-Host "  Start-PCSleep, Restart-PC, Stop-PC"
} catch {
    Write-Warning "Command verification failed: `$_"
}
"@

$result = powershell -NoProfile -Command $testScript

Write-Host "Setup completed successfully!" -ForegroundColor Green
Write-Host @"

PCPow has been installed:
1. PowerShell module installed to: $userModulePath
2. Scripts installed to: $shortcutsDir
3. Profile updated at: $PROFILE

You can now use the following commands from anywhere:
1. From Run menu (Win+R) or Command Prompt:
   pcpow sleep
   pcpow restart
   pcpow shutdown

2. From PowerShell:
   Quick aliases:
   pows            # Sleep
   powr            # Restart
   powd            # Shutdown
   
   Full commands:
   Start-PCSleep
   Restart-PC
   Stop-PC

Add -Force to any command to skip confirmation and force close apps.
Example: pcpow sleep -Force or pows -Force

Please close and reopen your PowerShell window for the changes to take effect.
"@

# Try to import the module immediately to verify installation
try {
    Import-Module $moduleName -Force -ErrorAction Stop
    Write-Host "`nModule successfully imported!" -ForegroundColor Green
} catch {
    Write-Host "`nWarning: Could not import module immediately. Error: $_" -ForegroundColor Yellow
    Write-Host "This is expected if running from a restricted execution policy." -ForegroundColor Yellow
}

# Prompt to restart PowerShell
$restart = Read-Host "Would you like to restart PowerShell now? (Y/N)"
if ($restart -eq 'Y' -or $restart -eq 'y') {
    Start-Process powershell
    exit
}

# Add these verification steps before the final message
Write-Host "Verifying installation..."

# Verify files and content
Write-Host "Checking installed files..."
$verificationErrors = @()

# Verify batch file and its content
if (-not (Test-Path "$shortcutsDir\pcpow.bat")) {
    $verificationErrors += "pcpow.bat not found in $shortcutsDir"
} else {
    $batchContent = Get-Content "$shortcutsDir\pcpow.bat" -Raw
    if (-not $batchContent.Contains("Start-PCSleep")) {
        $verificationErrors += "pcpow.bat help text is outdated"
    }
}

# Verify PowerShell scripts
$requiredScripts = @(
    "Close-AndSleep.ps1",
    "Close-AndRestart.ps1",
    "Close-AndShutdown.ps1"
)

foreach ($script in $requiredScripts) {
    if (-not (Test-Path "$shortcutsDir\$script")) {
        $verificationErrors += "$script not found in $shortcutsDir"
    }
}

# Verify module installation
if (-not (Get-Module -ListAvailable -Name pcpow-common)) {
    $verificationErrors += "pcpow-common module not found in PowerShell modules"
}

# Test PATH environment
$paths = $env:Path -split ';'
if (-not ($paths -contains $shortcutsDir)) {
    Write-Warning "$shortcutsDir is not in PATH. Commands may not work from Run prompt."
    Write-Host "Adding $shortcutsDir to PATH..."
    $env:Path = "$env:Path;$shortcutsDir"
    [Environment]::SetEnvironmentVariable(
        "Path",
        $env:Path,
        [System.EnvironmentVariableTarget]::User
    )
}

# Add the module path to PSModulePath if not already present
$modulePath = Split-Path -Parent $userModulePath
if (-not ($env:PSModulePath -split ';' -contains $modulePath)) {
    [Environment]::SetEnvironmentVariable(
        "PSModulePath",
        "$env:PSModulePath;$modulePath",
        [System.EnvironmentVariableTarget]::User
    )
}

# Add WindowsApps to PATH if not present
$windowsAppsPath = Join-Path $env:USERPROFILE "AppData\Local\Microsoft\WindowsApps"
$currentPath = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::User)
if ($currentPath -notlike "*$windowsAppsPath*") {
    [Environment]::SetEnvironmentVariable(
        "PATH",
        "$currentPath;$windowsAppsPath",
        [EnvironmentVariableTarget]::User
    )
}

# Report verification results
if ($verificationErrors.Count -gt 0) {
    Write-Warning "Installation verification found issues:"
    foreach ($error in $verificationErrors) {
        Write-Warning "  - $error"
    }
    Write-Warning "Please run the setup script again or check the documentation for troubleshooting."
} else {
    Write-Host "Installation verified successfully!" -ForegroundColor Green
} 