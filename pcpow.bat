@echo off
setlocal enabledelayedexpansion

REM Get the directory where the batch file is located
set "PCPOW_DIR=%~dp0"

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
powershell -ExecutionPolicy Bypass -NoProfile -File "%PCPOW_DIR%Close-AndSleep.ps1" %2
set EXIT_CODE=%errorlevel%
if %EXIT_CODE% NEQ 0 (
    echo Operation cancelled or failed with code %EXIT_CODE%
    exit /b %EXIT_CODE%
)
exit /b 0

:restart
powershell -ExecutionPolicy Bypass -NoProfile -File "%PCPOW_DIR%Close-AndRestart.ps1" %2
set EXIT_CODE=%errorlevel%
if %EXIT_CODE% NEQ 0 (
    echo Operation cancelled or failed with code %EXIT_CODE%
    exit /b %EXIT_CODE%
)
exit /b 0

:shutdown
powershell -ExecutionPolicy Bypass -NoProfile -File "%PCPOW_DIR%Close-AndShutdown.ps1" %2
set EXIT_CODE=%errorlevel%
if %EXIT_CODE% NEQ 0 (
    echo Operation cancelled or failed with code %EXIT_CODE%
    exit /b %EXIT_CODE%
)
exit /b 0

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
echo Configuration (pcpow.config.json):
echo   AlwaysForce: true/false  - Always run in force mode
echo   NoGraceful: true/false   - Skip graceful app closing
echo   timeoutMS: number        - Timeout for app closing (ms)
echo.
echo PowerShell Shortcuts:
echo   pows      - Sleep (alias for Sleep-PC)
echo   powr      - Restart (alias for Restart-PCApps)
echo   powd      - Shutdown (alias for Stop-PCApps)
echo.
echo Examples:
echo   pcpow sleep
echo   pcpow restart -Force
echo   pcpow shutdown
exit /b 0 