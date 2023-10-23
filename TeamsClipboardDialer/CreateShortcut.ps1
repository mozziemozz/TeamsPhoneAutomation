<#
    .SYNOPSIS
    This PowerShell script installs the Teams Clipboard Dialer app on your computer.

    .DESCRIPTION
    This script will copy the PowerShell script and the icon to the specified folder and create a shortcut for the script on your desktop.
    The shortcut has to be manually pinned to the taskbar. This step is part of the setup process when you launch 'createShortcut.ps1' for the first time.
    When the message appears, simply right click the blue phone icon to pin the app to the taskbar.

    .NOTES
    Author:             Martin Heusser | M365 Apps & Services MVP
    Version:            1.0.0
    Sponsor Project:    https://github.com/sponsors/mozziemozz
    Website:            https://heusser.pro

    .NOTES
    This script has been tested on Windows 11. It's assumed that your Desktop is redirected to OneDrive.
    The shortcut will be placed in $env:OneDrive\Desktop\TeamsClipboardDialer.lnk.
    The PowerShell script and the icon will be saved in $env:APPDATA\TeamsClipboardDialer.

    The phone icons are from https://icon-icons.com. Please see the 'IconAttribution.md' file for more information.

    .PARAMETER
    -InstallPath
        The path where the PowerShell script and the icon will be saved.
        Required:           false
        Type:               string
        Accepted values:    Any valid path
        Default value:      $env:APPDATA\TeamsClipboardDialer

    -Icon
        The icon that will be used for the shortcut.
        Required:           false
        Type:               string
        Accepted values:    1, 2
        Default value:      1

    -HotKey
        The hotkey that will be used to launch Teams Clipboard Dialer.
        Required:           false
        Type:               string
        Accepted values:    Any valid Shortcut combinations e.g. "CTRL+SHIFT+F8", $null
        Default value:      CTRL+SHIFT+F8

    .EXAMPLE
    Make sure to launch the script from the folder where it's located. Otherwise the script will not be able to copy the files to the specified folder.
    .\CreateShortcut.ps1
    .\CreateShortcut.ps1 -InstallPath "C:\Temp\TeamsClipboardDialer" -Icon "2"

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)][String]$InstallPath = "$env:APPDATA\TeamsClipboardDialer",
    [Parameter(Mandatory = $false)][ValidateSet("1", "2")][String]$Icon = "1",
    [Parameter(Mandatory = $false)][String]$HotKey = "CTRL+SHIFT+F8"

)

if (!(Test-Path -Path $InstallPath)) {

    New-Item -Path $InstallPath -ItemType Directory

}

switch ($Icon) {
    "1" {
        $phoneIcon = "PhoneIcon1.ico"
    }
    "2" {
        $phoneIcon = "PhoneIcon2.ico"
    }
    Default {
        $phoneIcon = "PhoneIcon1.ico"
    }
}

Copy-Item .\TeamsClipboardDialer.ps1 -Destination $InstallPath -Force
Copy-Item .\$phoneIcon -Destination $InstallPath -Force

# Credits: http://powershellblogger.com/2016/01/create-shortcuts-lnk-or-url-files-with-powershell/

$shell = New-Object -ComObject ("WScript.Shell")
$shortcut = $shell.CreateShortcut("$env:OneDrive\Desktop\TeamsClipboardDialer.lnk")

$shell = New-Object -ComObject ("WScript.Shell")
$shortcut = $shell.CreateShortcut("$env:OneDrive\Desktop\TeamsClipboardDialer.lnk")
$shortcut.TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe";
$shortcut.Arguments = "-ExecutionPolicy Bypass -File $InstallPath\TeamsClipboardDialer.ps1";
$shortcut.WorkingDirectory = "$InstallPath";
$shortcut.WindowStyle = 7;
$shortcut.Hotkey = "$HotKey";
$shortcut.IconLocation = "$InstallPath\$phoneIcon";
$shortcut.Description = "Teams Clipboard Dialer | https://heusser.pro";

$ShortCut.Save()

Set-Clipboard -Value "SetUpTeamsClipboardDialer"

Start-Process "$env:OneDrive\Desktop\TeamsClipboardDialer.lnk"

Start-Sleep -Seconds 5

Set-Clipboard -Value $null