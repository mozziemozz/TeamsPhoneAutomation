#Requires -Module Microsoft.Graph.Applications

Connect-MgGraph -Scopes "Application.ReadWrite.All"

$appDisplayName = "Read User Chats and Messages"
$folderPath = ".\Scripts\Teams\Chats"

$permissions = @(
    "Chat.Read.All",
    "ChatMessage.Read.All",
    "User.Read.All"
)

$graphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" # Microsoft Graph Service

$requiredPermissions = @()

foreach ($permission in $permissions) {

    $appRole = $graphServicePrincipal.AppRoles | Where-Object { $_.Value -eq $permission -and $_.AllowedMemberTypes -contains "Application" }

    if (-not $requiredPermissions.ResourceAppId) {
        $requiredPermissions += @{
            ResourceAppId = $graphServicePrincipal.AppId
            ResourceAccess = @()
        }
    }

    $requiredPermissions | Where-Object { $_.ResourceAppId -eq $graphServicePrincipal.AppId } | ForEach-Object {
        $_.ResourceAccess += @{
            Id = $appRole.Id
            Type = "Role"
        }
    }

}

$appRegistration = New-MgApplication -DisplayName $appDisplayName -SignInAudience "AzureADMyOrg" -Web @{ RedirectUris = @("https://localhost") } -RequiredResourceAccess $requiredPermissions

$servicePrincipal = New-MgServicePrincipal -AppId $appRegistration.AppId

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

Set-Content -Path "$folderPath\auth.json" -Value `
    (@{ appId = $appRegistration.AppId; tenantId = (Get-MgContext).TenantId; secret = $encryptedSecret } `
        | ConvertTo-Json) -Force -Encoding UTF8