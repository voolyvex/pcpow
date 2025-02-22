# Setup script to install shortcuts and aliases
$ErrorActionPreference = 'Stop'

# Get the directory where the scripts are located
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = [System.IO.Path]::GetFullPath($scriptPath)

# Create module directory in PowerShell modules path
$moduleName = "pcpow-common"
$moduleVersion = "1.0.0"
# Use PowerShell's built-in module path
$userModulePath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "WindowsPowerShell\Modules\$moduleName\$moduleVersion"

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
Write-Host "Copying $scriptPath\pcpow.config.json to $userModulePath"
Copy-Item "$scriptPath\pcpow.config.json" "$userModulePath\" -Force

# Verify module files
if (-not (Test-Path "$userModulePath\$moduleName.psm1")) {
    throw "Failed to copy module file to $userModulePath\$moduleName.psm1"
}
if (-not (Test-Path "$userModulePath\pcpow.config.json")) {
    throw "Failed to copy config file to $userModulePath\pcpow.config.json"
}
if (-not (Test-Path $manifestPath)) {
    throw "Failed to create module manifest at $manifestPath"
}

# Copy script files
Write-Host "Copying script files..."
Copy-Item "$scriptPath\pcpow.bat" $shortcutsDir -Force
Copy-Item "$scriptPath\Close-And*.ps1" $shortcutsDir -Force

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
function global:Sleep-PC {
    param([switch]`$Force)
    & "$shortcutsDir\Close-AndSleep.ps1" -Force:`$Force
}

function global:Restart-PCApps {
    param([switch]`$Force)
    & "$shortcutsDir\Close-AndRestart.ps1" -Force:`$Force
}

function global:Stop-PCApps {
    param([switch]`$Force)
    & "$shortcutsDir\Close-AndShutdown.ps1" -Force:`$Force
}

# Create aliases
Set-Alias -Name pows -Value Sleep-PC -Scope Global
Set-Alias -Name powr -Value Restart-PCApps -Scope Global
Set-Alias -Name powd -Value Stop-PCApps -Scope Global

Write-Host "PCPow commands loaded successfully" -ForegroundColor Green
# PCPow End
"@

Add-Content -Path $PROFILE -Value $moduleConfig

# Test the commands
Write-Host "`nTesting command availability..."
$testScript = @"
Import-Module $moduleName -Force
Get-Command -Name Sleep-PC, Restart-PCApps, Stop-PCApps, pows, powr, powd -ErrorAction SilentlyContinue
"@

$result = powershell -NoProfile -Command $testScript
if ($result) {
    Write-Host "Commands are available:" -ForegroundColor Green
    $result | ForEach-Object { Write-Host "  - $($_.Name)" }
} else {
    Write-Host "Warning: Commands not found. Please check the installation." -ForegroundColor Yellow
}

# Update pcpow.bat to use full paths
$batchContent = @"
@echo off
setlocal enabledelayedexpansion

if "%~1"=="" goto :help
if /i "%~1"=="sleep" goto :sleep
if /i "%~1"=="restart" goto :restart
if /i "%~1"=="shutdown" goto :shutdown
if /i "%~1"=="-h" goto :help
if /i "%~1"=="--help" goto :help

echo Error: Unknown command '%~1'
echo.
goto :help

:sleep
powershell -ExecutionPolicy Bypass -NoProfile -File "!shortcutsDir!\Close-AndSleep.ps1" %2
exit /b %errorlevel%

:restart
powershell -ExecutionPolicy Bypass -NoProfile -File "!shortcutsDir!\Close-AndRestart.ps1" %2
exit /b %errorlevel%

:shutdown
powershell -ExecutionPolicy Bypass -NoProfile -File "!shortcutsDir!\Close-AndShutdown.ps1" %2
exit /b %errorlevel%

:help
echo PCPow - Windows Power Management
echo -------------------------------
echo Usage: pcpow [command] [-Force]
echo.
echo Commands:
echo   sleep     - Close all apps and put PC to sleep
echo   restart   - Close all apps and restart PC
echo   shutdown  - Close all apps and shutdown PC
echo.
echo Options:
echo   -Force    - Skip confirmation and force close apps
echo.
echo Examples:
echo   pcpow sleep
echo   pcpow restart -Force
echo   pcpow shutdown
exit /b 0
"@

Set-Content -Path "$shortcutsDir\pcpow.bat" -Value $batchContent

Write-Host "Setup completed successfully!" -ForegroundColor Green
Write-Host @"

PCPow has been installed:
1. PowerShell module installed to: $userModulePath
2. Scripts installed to: $shortcutsDir
3. Profile updated at: $PROFILE

You can now use the following commands from anywhere:
1. From Run menu (Win+R) or Command Prompt:
   - pcpow sleep
   - pcpow restart
   - pcpow shutdown

2. From PowerShell:
   Short commands:
   - pow sleep   (or pows)
   - pow restart (or powr)
   - pow shutdown (or powd)
   
   Full commands:
   - Sleep-PC
   - Restart-PCApps
   - Stop-PCApps

Add -Force to any command to skip confirmation and force close apps.
Example: pow sleep -Force

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

# Verify batch file
if (-not (Test-Path "$shortcutsDir\pcpow.bat")) {
    Write-Warning "pcpow.bat not found in $shortcutsDir"
}

# Verify PowerShell scripts
$requiredScripts = @(
    "Close-AndSleep.ps1",
    "Close-AndRestart.ps1",
    "Close-AndShutdown.ps1"
)

foreach ($script in $requiredScripts) {
    if (-not (Test-Path "$shortcutsDir\$script")) {
        Write-Warning "$script not found in $shortcutsDir"
    }
}

# Verify module installation
if (-not (Get-Module -ListAvailable -Name pcpow-common)) {
    Write-Warning "pcpow-common module not found in PowerShell modules"
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