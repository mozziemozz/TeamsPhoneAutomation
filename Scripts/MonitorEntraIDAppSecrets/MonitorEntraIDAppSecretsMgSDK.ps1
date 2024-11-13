# Import external functions
. .\Modules\SecureCredsMgmt.ps1

$environment = Get-Content -Path .\Scripts\MonitorEntraIDAppSecrets\Environment.json | ConvertFrom-Json

$tenantId = $environment.TenantId
$appId = $environment.AppId
$appSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR((Get-Content -Path $environment.AppSecretFilePath | ConvertTo-SecureString))) | Out-String

$creds = New-Object System.Management.Automation.PSCredential ($appId, ($appSecret | ConvertTo-SecureString -AsPlainText -Force))

Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $creds -NoWelcome

$utcNow = (Get-Date).ToUniversalTime()

# Alert threshold in days
$alertThreshold = 30
$alertThresholdDateTime = $utcNow.AddDays(-$alertThreshold)

$allApplications = Get-MgApplication -All

foreach ($application in $allApplications) {

    if (!$application.passwordCredentials) {

        Write-Output "App Registration '$($application.displayName)' has no client secrets."
    }

    else {

        foreach ($passwordCredential in $application.passwordCredentials) {

            if (!$passwordCredential.displayName) {

                $passwordCredential.displayName = $passwordCredential.keyId
                
            }

            if ($passwordCredential.endDateTime -lt $alertThresholdDateTime) {

                if ($passwordCredential.endDateTime -lt $utcNow) {

                    Write-Output "Client secret '$($passwordCredential.displayName)' for app registration '$($application.displayName)' expired on '$($passwordCredential.endDateTime)'"

                }

                else {

                    Write-Output "Client secret '$($passwordCredential.displayName)' for app registration '$($application.displayName)' expires in less than $alertThreshold days on '$($passwordCredential.endDateTime)'"

                }

            }

            else {

                $validUntil = ($passwordCredential.endDateTime - $utcNow).Days

                Write-Output "Client secret '$($passwordCredential.displayName)' for app registration '$($application.displayName)' is valid for $validUntil more days until '$($passwordCredential.endDateTime)'"

            }

        }

    }

}