@echo off
setlocal enabledelayedexpansion

REM Get the directory where the batch file is located
set "PCPOW_DIR=%~dp0"
set "PCPOW_INSTALL_DIR=%USERPROFILE%\PCPow\bin"

if "%~1"=="" goto :help
if /i "%~1"=="sleep" goto :sleep
if /i "%~1"=="restart" goto :restart
if /i "%~1"=="shutdown" goto :shutdown
if /i "%~1"=="-h" goto :help
if /i "%~1"=="--help" goto :help
if /i "%~1"=="wake" goto :wake

echo Error: Unknown command '%~1'
echo.
goto :help

:sleep
powershell -ExecutionPolicy Bypass -NoProfile -File "%PCPOW_INSTALL_DIR%\PCPow-Launcher.ps1" -Action Sleep %2 %3
set EXIT_CODE=%errorlevel%
if %EXIT_CODE% NEQ 0 (
    echo Operation cancelled or failed with code %EXIT_CODE%
    exit /b %EXIT_CODE%
)
exit /b 0

:restart
powershell -ExecutionPolicy Bypass -NoProfile -File "%PCPOW_INSTALL_DIR%\PCPow-Launcher.ps1" -Action Restart %2 %3
set EXIT_CODE=%errorlevel%
if %EXIT_CODE% NEQ 0 (
    echo Operation cancelled or failed with code %EXIT_CODE%
    exit /b %EXIT_CODE%
)
exit /b 0

:shutdown
powershell -ExecutionPolicy Bypass -NoProfile -File "%PCPOW_INSTALL_DIR%\PCPow-Launcher.ps1" -Action Shutdown %2 %3
set EXIT_CODE=%errorlevel%
if %EXIT_CODE% NEQ 0 (
    echo Operation cancelled or failed with code %EXIT_CODE%
    exit /b %EXIT_CODE%
)
exit /b 0

:wake
if "%~2"=="" (
    echo Error: MAC address required for wake command
    echo.
    goto :help
)
powershell -ExecutionPolicy Bypass -NoProfile -File "%PCPOW_INSTALL_DIR%\PCPow.ps1" -Action Wake -MACAddress "%~2"
set EXIT_CODE=%errorlevel%
if %EXIT_CODE% NEQ 0 (
    echo Operation cancelled or failed with code %EXIT_CODE%
    exit /b %EXIT_CODE%
)
exit /b 0

:help
echo PCPow - Power Control utility
echo.
echo Usage:
echo   pcpow sleep [options]    - Put PC to sleep
echo   pcpow restart [options]  - Restart PC
echo   pcpow shutdown [options] - Shut down PC
echo   pcpow wake MAC-ADDRESS   - Wake remote PC
echo.
echo Options:
echo   -force      - Skip countdown and force close applications
echo   -skipaction - Test mode without performing action
echo.
exit /b 0 