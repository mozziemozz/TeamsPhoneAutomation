# Import external functions
. .\Functions\Connect-MsTeamsServicePrincipal.ps1
. .\Functions\Connect-MgGraphHTTP.ps1
. .\Modules\SecureCredsMgmt.ps1

. Get-MZZTenantIdTxt
. Get-MZZAppIdTxt

. Get-MZZSecureCreds -FileName AppSecret
$AppSecret = $passwordDecrypted

. Connect-MgGraphHTTP -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret

$utcNow = (Get-Date).ToUniversalTime()

# Alert threshold in days
$alertThreshold = 30
$alertThresholdDateTime = $utcNow.AddDays(-$alertThreshold)

$allPages = @()

$allApps = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/applications" -Headers $Header -Method Get -ContentType "application/json")
$allPages += $allApps.value

if ($allApps.'@odata.nextLink') {

        do {

            $allApps = (Invoke-RestMethod -Uri $allApps.'@odata.nextLink' -Headers $Header -Method Get -ContentType "application/json")
            $allPages += $allApps.value

        } until (
            !$allApps.'@odata.nextLink'
        )
        
}

$allApps = $allPages

foreach ($app in $allApps) {

    if (!$app.passwordCredentials) {

        Write-Output "App Registration '$($app.displayName)' has no client secrets."
    }

    else {

        foreach ($passwordCredential in $app.passwordCredentials) {

            if (!$passwordCredential.displayName) {

                $passwordCredential.displayName = $passwordCredential.keyId
                
            }

            if ($passwordCredential.endDateTime -lt $alertThresholdDateTime) {

                if ($passwordCredential.endDateTime -lt $utcNow) {

                    Write-Output "Client secret '$($passwordCredential.displayName)' for app registration '$($app.displayName)' expired on '$($passwordCredential.endDateTime)'"

                }

                else {

                    Write-Output "Client secret '$($passwordCredential.displayName)' for app registration '$($app.displayName)' expires in less than $alertThreshold days on '$($passwordCredential.endDateTime)'"

                }

            }

            else {

                $validUntil = ($passwordCredential.endDateTime - $utcNow).Days

                Write-Output "Client secret '$($passwordCredential.displayName)' for app registration '$($app.displayName)' is valid for $validUntil more days until '$($passwordCredential.endDateTime)'"

            }

        }

    }

}