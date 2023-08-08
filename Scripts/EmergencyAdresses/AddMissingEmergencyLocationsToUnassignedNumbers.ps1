# Import external functions
. .\Functions\Connect-MsTeamsServicePrincipal.ps1
. .\Functions\Get-CsOnlineNumbers.ps1
. .\Modules\SecureCredsMgmt.ps1

. Get-MZZTenantIdTxt
. Get-MZZAppIdTxt

. Get-MZZSecureCreds -FileName AppSecret
$AppSecret = $passwordDecrypted

. Connect-MsTeamsServicePrincipal -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret

# Get CsOnline Numbers
$allCsOnlineNumbers = . Get-CsOnlineNumbers

$numbersWithoutEmergencyLocation = $allCsOnlineNumbers | Where-Object {!$_.LocationId -and $_.Capability -notcontains "ConferenceAssignment" -and $_.City -ne "Toll-Free"}

foreach ($number in $numbersWithoutEmergencyLocation) {

    Write-Host "Setting emergency location for number $number..." -ForegroundColor Cyan

    $matchingEmergencyLocationId = ($allCsOnlineNumbers | Where-Object {$_.City -eq $number.City -and $_.LocationId -ne $null}).LocationId[0]

    Set-CsPhoneNumberAssignment -PhoneNumber $number.TelephoneNumber -LocationId $matchingEmergencyLocationId

}

