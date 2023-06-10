# Import external functions
. .\Functions\Connect-MsTeamsServicePrincipal.ps1
. .\Functions\Connect-MgGraphHTTP.ps1

$TenantId = Get-Content -Path .\.local\TenantId.txt
$AppId = Get-Content -Path .\.local\AppId.txt
$AppSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR((Get-Content -Path .\.local\AppSecret.txt | ConvertTo-SecureString))) | Out-String

. Connect-MsTeamsServicePrincipal -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret
. Connect-MgGraphHTTP -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret

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
    Write-Host $pythonResult

}

