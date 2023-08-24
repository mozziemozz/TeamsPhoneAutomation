<#
    .SYNOPSIS
    Script to validate Teams phone number assignments using reverse number lookup.

    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.0
    Sponsor Project:    https://github.com/sponsors/mozziemozz
    Website:            https://heusser.pro

    This script contains 2 functions. The first one acquires a Teams token for the Skype API. The second one validates the assignment of a Teams phone number using reverse number lookup.

    .PARAMETER AdminUser
    Specifies the admin user. If not provided, the current username is used.

    .PARAMETER MFA
    Indicates whether multi-factor authentication (MFA) is enabled.

    .PARAMETER LineURI
    Specifies the LineURI to validate.

    .EXAMPLE
    Import functions by dot sourcing:
    . .\ValidateTeamsReverseNumberLookup.ps1 

    Execute function like this:
    . Test-MZZTeamsLineURIAssignment -LineURI "+1234567890" -AdminUser "user@domain.com" -MFA $true
    . Test-MZZTeamsLineURIAssignment -LineURI "+1234567890" -AdminUser "user@domain.com" -MFA $false

    Note: this API doesn't require any Teams admin permissions since it only validates the assignment of a LineURI to a Teams user which can also be done from the Teams client.

#>
function Get-MZZTeamsTokenForSkyeApi {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]
        $AdminUser = $env:USERNAME,

        [Parameter(Mandatory=$false)]
        [bool]
        $MFA
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

        . Get-MZZSecureCreds -FileName "$($AdminUser)_otpsecret"

        $teamsSkypeApiToken = Prompt-Credentials -ClientId $ClientId -Resource "https://api.spaces.skype.com" -Credentials $secureCreds -OTPSecretKey $passwordDecrypted

    }

    else {

        $teamsSkypeApiToken = Prompt-Credentials -ClientId $ClientId -Resource "https://api.spaces.skype.com" -Credentials $secureCreds

    }

    $accessToken = $teamsSkypeApiToken.access_token

    $Header = @{Authorization = "Bearer $accessToken"}

}

function Test-MZZTeamsLineURIAssignment {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $LineURI,

        [Parameter(Mandatory = $false)]
        [string]
        $AdminUser = $env:USERNAME,

        [Parameter(Mandatory = $true)]
        [bool]
        $MFA
    )

    $Error.Clear()

    $ErrorActionPreference = "SilentlyContinue"

    $reverseNumberLookup = Invoke-RestMethod -Method Get -Uri "https://teams.microsoft.com/api/mt/emea/beta/phone/numbers/$LineURI/teamsidentity" -Headers $Header
    
    if ($Error[0].Exception.Message -match "401") {

        Write-Warning -Message "The Teams token has expired. Acquiring a new one..."

        . Get-MZZTeamsTokenForSkyeApi -AdminUser $AdminUser -MFA $MFA

        . Test-MZZTeamsLineURIAssignment -LineURI $LineURI -AdminUser $AdminUser -MFA $MFA

    }

    if ($Error[0].Exception.Message -match "404") {

        Write-Warning -Message "The LineURI $LineURI is not assigned to a Teams user or there is an issue with the assignment in."

    }

    else {
        
        Write-Host "The LineURI $LineURI is assigned to user $($reverseNumberLookup.objectId). The assignment has been validated by reverse number lookup successfully." -ForegroundColor Green

    }

}