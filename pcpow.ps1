# PowerShell script version of pcpow command
param(
    [Parameter(Position=0)]
    [ValidateSet('sleep','restart','shutdown','help','-h','--help')]
    [string]$command = 'help',
    
    [Parameter()]
    [switch]$Force
)

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

switch ($command) {
    "sleep" {
        & "$scriptPath\Close-AndSleep.ps1" -Force:$Force
    }
    "restart" {
        & "$scriptPath\Close-AndRestart.ps1" -Force:$Force
    }
    "shutdown" {
        & "$scriptPath\Close-AndShutdown.ps1" -Force:$Force
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
        Write-Host "Configuration (pcpow.config.json):"
        Write-Host "  AlwaysForce: true/false  - Skip confirmations"
        Write-Host "  NoGraceful: true/false   - Skip graceful closing"
        Write-Host "  timeoutMS: number        - Wait time for apps (ms)"
        Write-Host ""
        Write-Host "PowerShell Commands:"
        Write-Host "  Quick aliases:"
        Write-Host "  pows            - Sleep"
        Write-Host "  powr            - Restart"
        Write-Host "  powd            - Shutdown"
        Write-Host ""
        Write-Host "  Full commands:"
        Write-Host "  Start-PCSleep   - Sleep"
        Write-Host "  Restart-PC      - Restart"
        Write-Host "  Stop-PC         - Shutdown"
        Write-Host ""
        Write-Host "Examples:"
        Write-Host "  pcpow sleep"
        Write-Host "  pcpow restart -Force"
        Write-Host "  pcpow shutdown"
    }
} 