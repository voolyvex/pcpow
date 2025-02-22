@echo off
if "%~1"=="" goto :help
if /i "%~1"=="sleep" goto :sleep
if /i "%~1"=="restart" goto :restart
if /i "%~1"=="shutdown" goto :shutdown
if /i "%~1"=="-h" goto :help
if /i "%~1"=="--help" goto :help
goto :help

:sleep
powershell -ExecutionPolicy Bypass -File "%~dp0Close-AndSleep.ps1" %2
goto :eof

:restart
powershell -ExecutionPolicy Bypass -File "%~dp0Close-AndRestart.ps1" %2
goto :eof

:shutdown
powershell -ExecutionPolicy Bypass -File "%~dp0Close-AndShutdown.ps1" %2
goto :eof

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