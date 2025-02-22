$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

switch ($command) {
    "sleep" {
        & "$scriptPath\Close-AndSleep.ps1" $Force
    }
    "restart" {
        & "$scriptPath\Close-AndRestart.ps1" $Force
    }
    "shutdown" {
        & "$scriptPath\Close-AndShutdown.ps1" $Force
    }
    default {
        Write-Host "PCPow - Windows Power Management"
        Write-Host "-------------------------------"
        Write-Host "Usage: pcpow [command] [-Force]"
        Write-Host ""
        Write-Host "Commands:"
        Write-Host "  sleep     - Close all apps and put PC to sleep"
        Write-Host "  restart   - Close all apps and restart PC" 
        Write-Host "  shutdown  - Close all apps and shutdown PC"
        Write-Host ""
        Write-Host "Options:"
        Write-Host "  -Force    - Skip confirmation and force close apps"
        Write-Host ""
        Write-Host "Examples:"
        Write-Host "  pcpow sleep"
        Write-Host "  pcpow restart -Force"
        Write-Host "  pcpow shutdown"
    }
} 