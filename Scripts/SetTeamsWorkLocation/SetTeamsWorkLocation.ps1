<#
    .SYNOPSIS
    Script to set work location in Microsoft Teams.

    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.1
    Sponsor Project:    https://github.com/sponsors/mozziemozz
    Website:            https://heusser.pro

    This script sets the work location in Microsoft Teams based on the provided IP address and IP type (Local or Public). It requires the AADInternals module and external functions from SecureCredsMgmt.ps1.

    .PARAMETER AdminUser
    Specifies the admin user. If not provided, the current username is used.

    .PARAMETER MFA
    Indicates whether multi-factor authentication (MFA) is enabled.

    .PARAMETER IpType
    Specifies the IP type. Valid values are 'Local' and 'Public'.

    .PARAMETER IpAddress
    Specifies the IP address of your home/remote network.

    .EXAMPLE
    .\Set-TeamsWorkLocation.ps1 -AdminUser "user@domain.com" -MFA -IpType "Local" -IpAddress "192.168.1.10"

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]
    $AdminUser = $env:USERNAME,

    [Parameter(Mandatory=$false)]
    [switch]
    $MFA,

    [Parameter(Mandatory=$true)]
    [ValidateSet('Local', 'Public')]
    [string]
    $IpType,

    [Parameter(Mandatory=$true)]
    [string]
    $IpAddress
)

#requires -Modules "AADInternals"

# Import external functions
. .\Modules\SecureCredsMgmt.ps1

. Get-MZZSecureCreds -AdminUser $AdminUser

$functions = Get-ChildItem -Path "C:\Program Files\WindowsPowerShell\Modules\AADInternals\0.9.0" -Filter "*.ps1"

$functions = $functions | Where-Object {$_.Name -match "Teams" -or $_.Name -match "AccessToken" -or $_.Name -match "CommonUtils"}

foreach ($function in $functions) {

    . $function.FullName

}

# Teams Client Application Id
$ClientId = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"

if ($MFA) {

    . Get-MZZSecureCreds -FileName "otpsecret"

    $teamsPresenceToken = Prompt-Credentials -ClientId $ClientId -Resource "https://presence.teams.microsoft.com" -Credentials $secureCreds -OTPSecretKey $passwordDecrypted

}

else {

    # Alternative resource:
    # $teamsPresenceToken = Prompt-Credentials -ClientId $ClientId -Resource "https://api.spaces.skype.com" -Credentials $credentials

    $teamsPresenceToken = Prompt-Credentials -ClientId $ClientId -Resource "https://presence.teams.microsoft.com" -Credentials $secureCreds

}

$accessToken = $teamsPresenceToken.access_token

$Header = @{Authorization = "Bearer $accessToken"}

switch ($IpType) {
    Local {

        $netAdapter = (Get-NetAdapter | Where-Object {$_.Status -eq "Up"}).Name
        $netAddress = (Get-NetIPAddress | Where-Object {$_.InterfaceAlias -eq $netAdapter -and $_.AddressFamily -eq "IPv4"}).IPAddress

        $netAddressMatch = $netAddress.Split(".")[0] + "." + $netAddress.Split(".")[1] + "." + $netAddress.Split(".")[2]

        if ($IpAddress.StartsWith($netAddressMatch)) {

            $workLocation = 2
            $workLocationName = "Remote"

        }

        else {

            $workLocation = 1
            $workLocationName = "Office"

        }

    }
    Public {

        $ipInfo = Invoke-RestMethod -Method Get -Uri "https://ipapi.co/json"

        $IpAddressMatch = $IpAddress.Split(".")[0] + "." + $IpAddress.Split(".")[1] + "." + $IpAddress.Split(".")[2]

        if ($ipInfo.ip.StartsWith("$IpAddressMatch")) {

            $workLocation = 2
            $workLocationName = "Remote"

        }

        else {

            $workLocation = 1
            $workLocationName = "Office"

        }

    }
    Default {}
}


$endOfDay = ((Get-Date).ToUniversalTime().Date.AddDays(1).AddSeconds(-1).GetDateTimeFormats() | Where-Object {$_ -match "GMT"})[-1]

$body = @"
{
    "location": $workLocation,
    "expirationTime": "$endOfDay"
}
"@

Invoke-RestMethod -Method Put -Headers $Header -ContentType "application/json" -Body $body -Uri "https://presence.teams.microsoft.com/v1/me/workLocation"

if ($?) {

    Write-Host "Work location set to $workLocationName" -ForegroundColor Green

}

else {

    Write-Host "Error while trying to set work location to $workLocationName" -ForegroundColor Red

}