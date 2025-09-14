#Requires -Module Microsoft.Graph.Applications

Connect-MgGraph -Scopes "Application.ReadWrite.All"

$appDisplayName = "Call Queue Missed Call Reporting v2"

$appRegistration = New-MgApplication -DisplayName $appDisplayName -SignInAudience "AzureADMyOrg" -Web @{ RedirectUris = @("https://localhost") } -RequiredResourceAccess @(
    @{
        ResourceAppId = "00000003-0000-0000-c000-000000000000" # Microsoft Graph
        ResourceAccess = @(
            @{
                Id = "df021288-bdef-4463-88db-98f22de89214" # User.Read.All
                Type = "Role"
            }
            @{
                Id = "a2611786-80b3-417e-adaa-707d4261a5f0" # CallRecord-PstnCalls.Read.All
                Type = "Role"
            }
            @{
                Id = "45bbb07e-7321-4fd7-a8f6-3ff27e6a81c8	" # CallRecords.Read.All
                Type = "Role"
            }
        )
    }
)

$servicePrincipal = New-MgServicePrincipal -AppId $appRegistration.AppId

$graphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" # Microsoft Graph Service

foreach ($permission in $appRegistration.RequiredResourceAccess.ResourceAccess) {

    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id -PrincipalId $servicePrincipal.Id `
        -ResourceId $graphServicePrincipal.Id `
        -AppRoleId $permission.Id

}

$clientSecret = Add-MgApplicationPassword -ApplicationId $appRegistration.Id -PasswordCredential @{
    DisplayName = "Graph PowerShell"
    EndDateTime = (Get-Date).AddMonths(3)
}

$encryptedSecret = $clientSecret.SecretText | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString

Set-Content -Path ".\Scripts\Teams\CallRecords\auth.json" -Value `
    (@{ appId = $appRegistration.AppId; tenantId = (Get-MgContext).TenantId; secret = $encryptedSecret } `
        | ConvertTo-Json) -Force -Encoding UTF8