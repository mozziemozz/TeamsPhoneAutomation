# Import external functions
. .\Functions\Connect-MsTeamsServicePrincipal.ps1
. .\Functions\Connect-MgGraphHTTP.ps1
. .\Modules\SecureCredsMgmt.ps1

. Get-MZZTenantIdTxt
. Get-MZZAppIdTxt

. Get-MZZSecureCreds -FileName AppSecret
$AppSecret = $passwordDecrypted

. Connect-MgGraphHTTP -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret

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