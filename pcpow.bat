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
exit /b %errorlevel%

:restart
powershell -ExecutionPolicy Bypass -NoProfile -File "%PCPOW_DIR%Close-AndRestart.ps1" %2
exit /b %errorlevel%

:shutdown
powershell -ExecutionPolicy Bypass -NoProfile -File "%PCPOW_DIR%Close-AndShutdown.ps1" %2
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