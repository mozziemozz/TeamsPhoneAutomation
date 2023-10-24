<#
    .SYNOPSIS
    This PowerShell script allows you to dial a phone number from the clipboard using Microsoft Teams.

    .DESCRIPTION
    This script extracts a phone number from the clipboard and applies some normalization rules and other formatting operations such as trimming, removing special characters and invalid zeros.
    Unfortunately, many companies do not adhere to the E.164 standard and list their numbers in incorrect ways like e.g. +41 (0) 44 123 45 67. This format simply does not exist.
    The script uses Ken Lasko's awesome regular expression to remove any invalid zeros from the phone number before the number is called using Microsoft Teams.

    .NOTES
    Author:             Martin Heusser | M365 Apps & Services MVP
    Version:            1.0.4
    Changes:            2023-10-24
                        Add hint if clipboard is empty
                        Improve regex matching, add code to change 00 to +
                        Fix 00 to + replacement, add comments
                        Switch to BalloonTip instead of MessageBox for invalid clipboard content
    Sponsor Project:    https://github.com/sponsors/mozziemozz
    Website:            https://heusser.pro

    .NOTES
    This script has been tested on Windows 11 and the new Teams (2.1).
    The shortcut will be placed in $env:OneDrive\Desktop\TeamsClipboardDialer.lnk.

    The phone icons are from https://icon-icons.com. Please see the 'IconAttribution.md' file for more information.

    .EXAMPLE
    This script is not intended to be launched from PowerShell directly. Instead, use the file 'CreateShortcut.ps1' to create a shortcut on your desktop and pin it to the taskbar.
    The shortcut has to be manually pinned to the taskbar. This step is part of the setup process when you launch 'createShortcut.ps1' for the first time.

#>

$phoneNumber = Get-Clipboard | Out-String

$originalClipboardValue = $phoneNumber.Trim()

# Trim eventual whitespaces
$phoneNumber = $phoneNumber.Trim()

# Remove any non-digit or + characters
$phoneNumber = $phoneNumber -replace '[^\d\+]'

# Replace leading 00 with +
$phoneNumber = $phoneNumber -replace '^00', '+'

# Remove any invalid zeros
# Credits: https://ucken.blogspot.com/2016/03/trunk-prefixes-in-skype4b.html
$phoneNumber = $phoneNumber -replace ('^\+(1|7|2[07]|3[0-46]|39\d|4[013-9]|5[1-8]|6[0-6]|8[1246]|9[0-58]|2[1235689]\d|24[013-9]|242\d|3[578]\d|42|5[09]\d|6[789]\d|8[035789]\d|9[679]\d)(?:0)?(\d{6,14})?$', '+$1$2')

# Check if there is a phone number in the clipboard
if ($phoneNumber -notmatch '^(?:\+\d+|\d+)') {

    if ($originalClipboardValue -eq "SetUpTeamsClipboardDialer") {

        $Message = "Right-click on the blue phone icon in the taskbar`nand select 'Pin to taskbar' to pin the app.`n`nClick OK when you're done."
        $Title = "Teams Clipboard Dialer | SETUP"

        # Show hint if there is no phone number in clipboard
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        [Windows.Forms.MessageBox]::Show($Message, $Title, [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)
        
    }

    else {

        # Check if $phoneNumber is null or whitespace
        if (!$originalClipboardValue) {

            $clipboardContent = "Clipboard is empty."

        }

        else {

            # Get the clipboard content
            $clipboardContent = $originalClipboardValue

        }

        # $Message = "Clipboard doesn't contain a phone number. Copy a phone number and try again. Clipboard content: $clipboardContent"
        $Message = @"
Clipboard doesn't contain a phone number. 
Copy a phone number and try again. 

Clipboard content: $clipboardContent

"@
        $Title = "Teams Clipboard Dialer"

        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        $balloonTip = New-Object System.Windows.Forms.NotifyIcon
        $balloonTip.Icon = [System.Drawing.SystemIcons]::Information
        $balloonTip.BalloonTipIcon = "Info"
        $balloonTip.BalloonTipTitle = $Title
        $balloonTip.BalloonTipText = $Message 
        $balloonTip.Visible = $True
        $balloonTip.ShowBalloonTip(30000)

    }
    
}

else {

    # Replace + with %2B (url encoding)
    $phoneNumber = $phoneNumber.Replace("+", "%2B")

    # Launch Teams to dial the number
    Start-Process ms-teams "https://teams.microsoft.com/l/call/0/0?users=4:$phoneNumber"

}