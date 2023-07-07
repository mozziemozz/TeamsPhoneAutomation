#requires -Modules "AADInternals"

# Import external functions
. .\Modules\SecureCredsMgmt.ps1

. Get-MZZSecureCreds -AdminUser ""
. Get-MZZSecureCreds -FileName "otpsecret"

$functions = Get-ChildItem -Path "C:\Program Files\WindowsPowerShell\Modules\AADInternals\0.9.0" -Filter "*.ps1"
# $functions = $functions | Where-Object {$_.Name -notmatch "AADSyncSettings"}

$functions = $functions | Where-Object {$_.Name -match "Teams" -or $_.Name -match "AccessToken" -or $_.Name -match "CommonUtils"}

foreach ($function in $functions) {

    . $function.FullName

}

# Teams Client
$ClientId = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"

$teamsPresenceToken = Prompt-Credentials -ClientId $ClientId -Resource "https://presence.teams.microsoft.com" -Credentials $secureCreds -OTPSecretKey $passwordDecrypted
# $teamsPresenceToken = Prompt-Credentials -ClientId $ClientId -Resource "https://api.spaces.skype.com" -Credentials $credentials -OTPSecretKey ""


$accessToken = $teamsPresenceToken.access_token

$Header = @{Authorization = "Bearer $accessToken"}

$ipInfo = Invoke-RestMethod -Method Get -Uri "https://ipapi.co/json"

if ($ipInfo.ip.StartsWith("87.241.34.")) {

    $workLocation = 1
    $workLocationName = "Office"

}

else {

    $workLocation = 2
    $workLocationName = "Remote"

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
