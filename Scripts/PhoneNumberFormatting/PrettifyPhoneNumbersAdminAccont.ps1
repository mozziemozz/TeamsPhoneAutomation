<#
    .SYNOPSIS
    Script to prettify phone numbers for Teams users and update them in Azure Active Directory.

    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.0
    Sponsor Project:    https://github.com/sponsors/mozziemozz
    Website:            https://heusser.pro
    Link:               https://github.com/mozziemozz/TeamsPhoneAutomation/blob/main/Scripts/PhoneNumberFormatting/PrettifyPhoneNumbersServicePrincipal.ps1

    This script connects to Microsoft Teams and Azure Active Directory using service principal credentials and retrieves Teams users with phone numbers assigned. It then prettifies the phone numbers using the `phonenumbers` Python library and updates the formatted numbers in Azure Active Directory for each user.

    .NOTES
    - Requires Python to be installed
    - Requires "phonenumbers" python library

#>

Invoke-Expression (Invoke-RestMethod -Method Get -Uri "https://raw.githubusercontent.com/mozziemozz/M365CallFlowVisualizer/main/Functions/Connect-M365CFV.ps1")

. Connect-M365CFV

# Get all Teams users which have a phone number assigned
$allTeamsPhoneUsers = Get-CsOnlineUser -Filter "LineURI -ne `$null"

foreach ($teamsPhoneUser in $allTeamsPhoneUsers) {

    $phoneNumber = $teamsPhoneUser.LineURI.Replace("tel:","")

    $pythonCode = @"
import phonenumbers

def format_phone_number(phone_number):
    parsed_number = phonenumbers.parse(phone_number, None)
    if phonenumbers.is_valid_number(parsed_number):
        formatting_pattern = phonenumbers.format_number(parsed_number, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
        return formatting_pattern
    else:
        return "Invalid phone number"

phone_number = "$phoneNumber"
formatting_pattern = format_phone_number(phone_number)
print(formatting_pattern)
"@
    
    $pythonResult = python -c $pythonCode

    $pythonResult = $pythonResult.Replace("-"," ")

    Update-MgUser -UserId $teamsPhoneUser.Identity -BusinessPhones @($pythonResult)

    Write-Host "LineURI $($teamsPhoneUser.LineUri) has been prettifyied to $pythonResult and set in AAD for user $($teamsPhoneUser.UserPrincipalName)"

}