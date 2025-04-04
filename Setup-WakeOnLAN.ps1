# Requires -Version 5.1
<#
.SYNOPSIS
Configures Wake-on-LAN on the current PC

.DESCRIPTION
Configures network adapters to allow wake from shutdown and remote wake-on-LAN

.PARAMETER AllowRemoteAccess
If specified, configures Windows to allow remote wake access from FREIA laptop

.PARAMETER FreiaIP
The IPv6 address of the FREIA laptop for remote access configuration

.EXAMPLE
.\Setup-WakeOnLAN.ps1
Sets up Wake-on-LAN with standard settings

.EXAMPLE
.\Setup-WakeOnLAN.ps1 -AllowRemoteAccess
Sets up Wake-on-LAN and configures remote access rules for FREIA
#>

[CmdletBinding()]
param (
    [switch]$AllowRemoteAccess,
    [string]$FreiaIP = "fe80::d698:81ea:a618:7b4c%17"
)

Write-Host "Wake-on-LAN Configuration Tool" -ForegroundColor Cyan
Write-Host "-----------------------------" -ForegroundColor Cyan
Write-Host ""

# Check if running as admin
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "This script requires administrator privileges." -ForegroundColor Red
    Write-Host "Please restart the script with administrative rights." -ForegroundColor Yellow
    exit 1
}

# Step 1: Configure network adapters to enable Wake-on-LAN
Write-Host "Step 1: Configuring network adapters for Wake-on-LAN..." -ForegroundColor Green

$adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
Write-Host "Found $($adapters.Count) active network adapters" -ForegroundColor White

foreach ($adapter in $adapters) {
    Write-Host " - Configuring adapter: $($adapter.Name) ($($adapter.InterfaceDescription))" -ForegroundColor Yellow
    
    # Try to enable Wake-on-LAN in adapter settings
    try {
        # Check if the adapter supports Wake-on-LAN
        $supportWol = $null
        try {
            $supportWol = Get-NetAdapterAdvancedProperty -Name $adapter.Name -RegistryKeyword "*WakeOnMagicPacket*" -ErrorAction SilentlyContinue
        } catch {}
         
        if ($supportWol) {
            Write-Host "   Setting Wake on Magic Packet: Enabled" -ForegroundColor White
            Set-NetAdapterAdvancedProperty -Name $adapter.Name -RegistryKeyword "*WakeOnMagicPacket*" -RegistryValue "1" -ErrorAction SilentlyContinue
        }
        
        # Some adapters use different keywords
        $supportWol = $null
        try {
            $supportWol = Get-NetAdapterAdvancedProperty -Name $adapter.Name -RegistryKeyword "*WakeOn*" -ErrorAction SilentlyContinue
        } catch {}
        
        if ($supportWol) {
            Write-Host "   Setting Wake on Magic Packet (WakeOn): Enabled" -ForegroundColor White
            Set-NetAdapterAdvancedProperty -Name $adapter.Name -RegistryKeyword "*WakeOn*" -RegistryValue "6" -ErrorAction SilentlyContinue
        }
        
        # Try additional common registry keywords for different network cards
        try {
            $wolPatterns = @("*WakeMagicPacket*", "*PMWakeOnMagicPacket*", "*WakeOnPattern*")
            foreach ($pattern in $wolPatterns) {
                $wolProperty = Get-NetAdapterAdvancedProperty -Name $adapter.Name -RegistryKeyword $pattern -ErrorAction SilentlyContinue
                if ($wolProperty) {
                    Write-Host "   Setting ${pattern}: Enabled" -ForegroundColor White
                    Set-NetAdapterAdvancedProperty -Name $adapter.Name -RegistryKeyword $pattern -RegistryValue "1" -ErrorAction SilentlyContinue
                }
            }
        } catch {}
        
        # Disable the "Allow the computer to turn off this device to save power" option
        Write-Host "   Configuring power management settings" -ForegroundColor White
        $adapterInstance = Get-WmiObject -Class Win32_NetworkAdapter | Where-Object { $_.DeviceID -eq $adapter.DeviceID }
        if ($adapterInstance) {
            $pnpInstance = Get-WmiObject -Class Win32_PnPEntity | Where-Object { $_.PNPDeviceID -eq $adapterInstance.PNPDeviceID }
            if ($pnpInstance) {
                $devicePowerMgmt = Get-WmiObject -Class MSPower_DeviceEnable -Namespace root\wmi | Where-Object { $_.InstanceName -like "*$($pnpInstance.PNPDeviceID)*" }
                if ($devicePowerMgmt) {
                    $devicePowerMgmt.Enable = $false
                    $devicePowerMgmt.Put() | Out-Null
                    Write-Host "   Disabled power saving for this adapter" -ForegroundColor White
                }
            }
        }
        
        # Enable Wake-on-LAN capabilities
        Write-Host "   Enabling Wake-on-LAN power settings" -ForegroundColor White
        powercfg /DEVICEENABLEWAKE "$($adapter.PnPDeviceID)" 2>$null
        
        Write-Host "   Configuration completed for this adapter" -ForegroundColor Green
    } catch {
        Write-Host "   Error configuring adapter: $_" -ForegroundColor Red
    }
}

# Step 2: Ensure the system is configured to wake from shutdown
Write-Host "`nStep 2: Configuring system power settings..." -ForegroundColor Green

# Enable hybrid sleep
powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_SLEEP HYBRIDSLEEP 1
powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_SLEEP HYBRIDSLEEP 1

# Allow wake timers
powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_SLEEP RTCWAKE 1
powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_SLEEP RTCWAKE 1

# Ensure Fast Startup is disabled (can interfere with WoL)
Write-Host " - Disabling Fast Startup (can interfere with Wake-on-LAN)..." -ForegroundColor Yellow
$powerRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
Set-ItemProperty -Path $powerRegPath -Name HiberbootEnabled -Value 0 -Type DWord -Force

# Apply changes
powercfg /SETACTIVE SCHEME_CURRENT

# Step 3: Configure Windows Firewall for WoL (if requested)
if ($AllowRemoteAccess) {
    Write-Host "`nStep 3: Configuring firewall for remote Wake-on-LAN access..." -ForegroundColor Green
    
    # Creating firewall rule to allow WoL traffic (UDP port 9)
    Write-Host " - Creating firewall rule for Wake-on-LAN traffic" -ForegroundColor Yellow
    
    # Remove existing rule if it exists
    Remove-NetFirewallRule -Name "Allow-WakeOnLAN" -ErrorAction SilentlyContinue
    
    # Create new rule
    New-NetFirewallRule -Name "Allow-WakeOnLAN" -DisplayName "Wake-on-LAN (UDP-In)" `
        -Description "Allows incoming Wake-on-LAN magic packets" `
        -Protocol UDP -LocalPort 9 -Action Allow -Enabled True `
        -Profile Any -Direction Inbound | Out-Null
    
    # Check if we have a specific IP for FREIA
    if (-not [string]::IsNullOrEmpty($FreiaIP)) {
        # Strip scope ID from IPv6 address if present
        $cleanIP = $FreiaIP -replace '%\d+$', ''
        Write-Host " - Creating specific rule for FREIA laptop ($cleanIP)" -ForegroundColor Yellow
        Remove-NetFirewallRule -Name "Allow-WakeOnLAN-FREIA" -ErrorAction SilentlyContinue
        New-NetFirewallRule -Name "Allow-WakeOnLAN-FREIA" -DisplayName "Wake-on-LAN from FREIA (UDP-In)" `
            -Description "Allows incoming Wake-on-LAN magic packets from FREIA laptop" `
            -Protocol UDP -LocalPort 9 -Action Allow -Enabled True `
            -RemoteAddress $cleanIP `
            -Profile Any -Direction Inbound | Out-Null
    }
    
    Write-Host " - Firewall rules created successfully" -ForegroundColor Green
}

# Step 4: Output MAC address information for using with wake commands
Write-Host "`nStep 4: Collecting MAC addresses for wake commands..." -ForegroundColor Green

$macAddresses = @()
foreach ($adapter in $adapters) {
    $macAddress = $adapter.MacAddress
    # Format MAC for WoL packet (replace - with :)
    $formattedMac = $macAddress -replace "-", ":"
    
    Write-Host " - Network adapter: $($adapter.Name)" -ForegroundColor Yellow
    Write-Host "   MAC Address: $formattedMac" -ForegroundColor White
    $macAddresses += [PSCustomObject]@{
        Name = $adapter.Name
        MAC = $formattedMac
        FormattedMAC = $formattedMac
        OriginalMAC = $macAddress
    }
}

# Step 5: Save MAC addresses to a file for future reference
$outputFile = Join-Path $PSScriptRoot "wake-targets.txt"
"# Wake-on-LAN Targets" | Out-File $outputFile
"# Generated on $(Get-Date)" | Out-File $outputFile -Append
"" | Out-File $outputFile -Append
"# This PC:" | Out-File $outputFile -Append
foreach ($mac in $macAddresses) {
    "$($mac.Name): $($mac.FormattedMAC)" | Out-File $outputFile -Append
}
"" | Out-File $outputFile -Append

# Add FREIA information if we have it
if ($AllowRemoteAccess) {
    "# FREIA Laptop:" | Out-File $outputFile -Append
    "# IP: $FreiaIP" | Out-File $outputFile -Append
    
    # Create a batch file for FREIA to wake this PC
    $freiaBatchFile = Join-Path $PSScriptRoot "wake-from-freia.bat"
    "@echo off" | Out-File $freiaBatchFile
    "echo Sending Wake-on-LAN packet to PC" | Out-File $freiaBatchFile -Append
    "echo MAC Address: $($macAddresses[0].FormattedMAC)" | Out-File $freiaBatchFile -Append
    "echo." | Out-File $freiaBatchFile -Append
    "powershell -ExecutionPolicy Bypass -NoProfile -Command ""& { `$mac='$($macAddresses[0].FormattedMAC)'; `$macByteArray=`$mac.Split(':','-') | ForEach-Object {[byte]('0x'+`$_)}; [byte[]]`$magicPacket = (,0xFF * 6) + (`$macByteArray * 16); `$udpClient = New-Object System.Net.Sockets.UdpClient; `$udpClient.Connect([System.Net.IPAddress]::Broadcast,9); `$udpClient.Send(`$magicPacket,`$magicPacket.Length); `$udpClient.Close(); Write-Host 'Wake-on-LAN packet sent to MAC: $($macAddresses[0].FormattedMAC)' }""" | Out-File $freiaBatchFile -Append
    "pause" | Out-File $freiaBatchFile -Append
    
    Write-Host " - Created wake-from-freia.bat file for FREIA to wake this PC" -ForegroundColor Green
}

Write-Host "`nConfiguration completed!" -ForegroundColor Cyan
Write-Host "MAC addresses saved to: $outputFile" -ForegroundColor Green
Write-Host ""
Write-Host "Important: Wake-on-LAN also requires BIOS/UEFI configuration!" -ForegroundColor Yellow
Write-Host "Please ensure 'Wake-on-LAN' or similar setting is enabled in your BIOS/UEFI." -ForegroundColor Yellow
Write-Host ""
Write-Host "To wake this PC remotely, use: pcpow wake [MAC-ADDRESS]" -ForegroundColor White
Write-Host "Example: pcpow wake $($macAddresses[0].FormattedMAC)" -ForegroundColor White
Write-Host ""

if ($AllowRemoteAccess) {
    Write-Host "Remote Access Information for FREIA laptop:" -ForegroundColor Cyan
    Write-Host " - This PC's MAC Addresses:" -ForegroundColor White
    foreach ($mac in $macAddresses) {
        Write-Host "   * $($mac.Name): $($mac.FormattedMAC)" -ForegroundColor White
    }
    Write-Host " - Firewall rules created:" -ForegroundColor White
    Write-Host "   * Allow-WakeOnLAN (all sources)" -ForegroundColor White
    Write-Host "   * Allow-WakeOnLAN-FREIA (specific to FREIA laptop)" -ForegroundColor White
    Write-Host "" 
    Write-Host " - Created wake-from-freia.bat for easy waking" -ForegroundColor White
    Write-Host ""
    Write-Host "From FREIA laptop, use the following command to wake this PC:" -ForegroundColor Yellow
    Write-Host "pcpow wake $($macAddresses[0].FormattedMAC)" -ForegroundColor White
}

Write-Host "`nReboot your PC to ensure all settings take effect." -ForegroundColor Magenta 