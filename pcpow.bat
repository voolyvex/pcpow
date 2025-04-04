@echo off
setlocal enabledelayedexpansion

REM Get the directory where the batch file is located
set "PCPOW_DIR=%~dp0"
cd /D "%PCPOW_DIR%"
if %ERRORLEVEL% neq 0 (
    echo Error: Failed to change to PCPow directory
    exit /b 1
)

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
powershell -ExecutionPolicy Bypass -NoProfile -File "%PCPOW_DIR%PCPow.ps1" -Action Sleep %2 %3
set EXIT_CODE=%errorlevel%
if %EXIT_CODE% NEQ 0 (
    echo Operation cancelled or failed with code %EXIT_CODE%
    exit /b %EXIT_CODE%
)
exit /b 0

:restart
powershell -ExecutionPolicy Bypass -NoProfile -File "%PCPOW_DIR%PCPow.ps1" -Action Restart %2 %3
set EXIT_CODE=%errorlevel%
if %EXIT_CODE% NEQ 0 (
    echo Operation cancelled or failed with code %EXIT_CODE%
    exit /b %EXIT_CODE%
)
exit /b 0

:shutdown
powershell -ExecutionPolicy Bypass -NoProfile -File "%PCPOW_DIR%PCPow.ps1" -Action Shutdown %2 %3
set EXIT_CODE=%errorlevel%
if %EXIT_CODE% NEQ 0 (
    echo Operation cancelled or failed with code %EXIT_CODE%
    exit /b %EXIT_CODE%
)
exit /b 0

:wake
if "%~2"=="" (
    echo Error: Target MAC address required
    echo Usage: pcpow wake [MAC_ADDRESS]
    exit /b 1
)
powershell -ExecutionPolicy Bypass -NoProfile -Command "& { $mac='%~2'; try { $formattedMac = & '%PCPOW_DIR%PCPow.ps1' -FormatMAC $mac } catch { Write-Error $_; exit 1 }; $macByteArray=$formattedMac.Split(':','-') | ForEach-Object {[byte]('0x'+$_)}; [byte[]]$magicPacket = (,0xFF * 6) + ($macByteArray * 16); $udpClient = New-Object System.Net.Sockets.UdpClient; $udpClient.Connect([System.Net.IPAddress]::Broadcast,9); $udpClient.Send($magicPacket,$magicPacket.Length); $udpClient.Close(); Write-Host 'Wake-on-LAN packet sent to MAC: $formattedMac' }"
set EXIT_CODE=%errorlevel%
if %EXIT_CODE% NEQ 0 (
    echo Sending WoL packet failed with code %EXIT_CODE%
    exit /b %EXIT_CODE%
)
exit /b 0

:help
echo PCPow - Windows Power Management
echo -------------------------------
echo Usage: pcpow [command] [options]
echo.
echo Commands:
echo   sleep     - Close all apps and put PC to sleep
echo   restart   - Close all apps and restart PC
echo   shutdown  - Close all apps and shutdown PC
echo   wake      - Send Wake-on-LAN packet to wake a remote PC
echo.
echo Options:
echo   -Force    - Skip confirmation and force close apps
echo   -SkipAction - Close apps but skip power action (testing)
echo.
echo Examples:
echo   pcpow sleep
echo   pcpow restart -Force
echo   pcpow shutdown
echo   pcpow wake 01:23:45:67:89:AB
echo.
echo Configuration (pcpow.config.json):
echo   AlwaysForce: true/false  - Skip confirmations
echo   NoGraceful: true/false   - Skip graceful closing
echo   TimeoutMS: number        - Wait time for apps (ms)
echo.
exit /b 0 