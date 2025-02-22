# Requires -Version 5.1
<#
.SYNOPSIS
Safely closes applications and shuts down the computer
.DESCRIPTION
Gracefully closes user applications before initiating system shutdown
#>

param([switch]$Force = $false)

Import-Module $PSScriptRoot\pcpow-common.psm1 -Force

if (-not $Force) {
    if (-not (Show-ConfirmationPrompt -ActionType "shutdown")) {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        exit
    }
}

Invoke-PowerAction -Action Shutdown -Force:$Force 