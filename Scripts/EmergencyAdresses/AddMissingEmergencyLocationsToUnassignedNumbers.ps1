<#

.SYNOPSIS
    This script sets emergency locations for unassigned phone numbers in Microsoft Teams.

.DESCRIPTION
    This script connects to the Microsoft Teams service using a service principal and sets
    emergency locations for phone numbers that do not have a location assigned.

.NOTES
    Author:             Martin Heusser
    Version:            1.0.0
    Prerequisites:      PowerShell with MicrosoftTeams module version 5.4.0,
                        Entra ID service principal with Teams Administrator rights
    Sponsor Project:    https://github.com/sponsors/mozziemozz
    Website:            https://heusser.pro
    Link:               https://github.com/mozziemozz/TeamsPhoneAutomation/blob/main/Scripts/EmergencyAdresses/AddMissingEmergencyLocationsToUnassignedNumbers.ps1

#>

# Requires minimum version of the MicrosoftTeams module
#Requires -Modules @{ ModuleName = "MicrosoftTeams"; ModuleVersion = "5.4.0" }

# Import external functions
. .\Functions\Connect-MsTeamsServicePrincipal.ps1
. .\Functions\Get-CsOnlineNumbers.ps1
. .\Modules\SecureCredsMgmt.ps1

# Get tenant ID and application ID from configuration files
. Get-MZZTenantIdTxt
. Get-MZZAppIdTxt

# Get secure credentials and decrypt the app secret
. Get-MZZSecureCreds -FileName AppSecret
$AppSecret = $passwordDecrypted

# Connect to Microsoft Teams service principal
. Connect-MsTeamsServicePrincipal -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret

# Get CsOnline Numbers
$allCsOnlineNumbers = . Get-CsOnlineNumbers

# Filter phone numbers without emergency location
$numbersWithoutEmergencyLocation = $allCsOnlineNumbers | Where-Object {!$_.LocationId -and $_.LocationUpdateSupported -eq $true}

# Iterate through numbers without emergency location
foreach ($number in $numbersWithoutEmergencyLocation) {
    Write-Host "Setting emergency location for number $number..." -ForegroundColor Cyan

    # Find matching emergency location ID
    $matchingEmergencyLocationId = ($allCsOnlineNumbers | Where-Object {$_.City -eq $number.City -and $_.LocationId -ne $null}).LocationId

    if ($matchingEmergencyLocationId.Count -gt 1) {

        $matchingEmergencyLocationId = $matchingEmergencyLocationId[0]

    }

    if ($matchingEmergencyLocationId) {

        Write-Host "Setting Location Id '$matchingEmergencyLocationId' for number $($number.TelephoneNumber)..."
        
        Set-CsPhoneNumberAssignment -PhoneNumber $number.TelephoneNumber -LocationId $matchingEmergencyLocationId

    } 
    
    else {

        Write-Host "No existing Location Id found for number $($number.Telephonenumber)..."

    }

}
