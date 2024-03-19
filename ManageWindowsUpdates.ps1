# Check and elevate permissions if not already running as Administrator
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

function Disable-WindowsUpdate {
    # Stops and disables the Windows Update service
    Write-Host "Disabling Windows Update Service..."
    Stop-Service -Name wuauserv -Force
    Set-Service -Name wuauserv -StartupType Disabled
    Write-Host "Windows Update Service Disabled."
}

function Enable-WindowsUpdate {
    # Sets the Windows Update service to automatic and starts it
    Write-Host "Enabling Windows Update Service..."
    Set-Service -Name wuauserv -StartupType Automatic
    Start-Service -Name wuauserv
    Write-Host "Windows Update Service Enabled."
}

# Prompting the user
$userChoice = Read-Host "Would you like to disable or enable Windows Automatic Updates? (enable/disable)"
switch ($userChoice.ToLower()) {
    "disable" {
        Disable-WindowsUpdate
    }
    "enable" {
        Enable-WindowsUpdate
    }
    default {
        Write-Host "Invalid input. Please enter 'enable' or 'disable'."
    }
}
