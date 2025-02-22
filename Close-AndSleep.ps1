# Requires -Version 5.1
<#
.SYNOPSIS
Safely closes applications and puts computer to sleep
.DESCRIPTION
Gracefully closes user applications before initiating sleep mode
#>

param([switch]$Force = $false)

Import-Module $PSScriptRoot\pcpow-common.psm1 -Force

if (-not $Force) {
    if (-not (Show-ConfirmationPrompt -ActionType "sleep")) {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        exit
    }
}

Invoke-PowerAction -Action Sleep -Force:$Force 