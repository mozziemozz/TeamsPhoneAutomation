[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$managedIdentityObjectId,

    [Parameter(Mandatory = $false)]
    [string]$managedIdentityDisplayName,

    [Parameter(Mandatory = $true)]
    [array]$permissionsToAdd
)

# Connect to graph using delegated permissions
Connect-MgGraph -Scopes "Directory.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"

# Test
$managedIdentityObjectId = "583732bc-5cce-419b-a734-9d4ddb9977bb"
$managedIdentityDisplayName = ""

$permissionsToAdd = @(

    "Application.Read.All"
    # "Directory.Read.All"

)

if ($managedIdentityObjectId) {

    $managedIdentity = Get-MgServicePrincipal -ServicePrincipalId $managedIdentityObjectId


}

elseif ($managedIdentityDisplayName) {

    $managedIdentity = Get-MgServicePrincipal -Filter "DisplayName eq '$managedIdentityDisplayName'"

}

else {

    Write-Output "Please provide either the managed identity object ID or display name."

    exit

}

$graphServicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"

$existingPermissions = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $managedIdentity.Id

foreach ($permission in $permissionsToAdd) {

    # Find app role with those permissions

    $graphPermissionToAdd = $graphServicePrincipal.AppRoles | Where-Object { $_.AllowedMemberTypes -contains "Application" -and $_.Value -eq $permission }

    if ($existingPermissions.AppRoleId -notcontains $graphPermissionToAdd.Id) {

        Write-Output "Adding permission '$permission' to managed identity '$($managedIdentity.DisplayName)'"

        $bodyParam = @{
            PrincipalId = $managedIdentity.Id # This is the object id of the managed identity / service principal, not the app id
            ResourceId  = $graphServicePrincipal.Id # This is the object id of the graph service principal not the app id: 00000003-0000-0000-c000-000000000000
            AppRoleId   = $graphPermissionToAdd.Id # This is the permission id: https://learn.microsoft.com/en-us/graph/migrate-azure-ad-graph-permissions-differences
        }

        # Assign permission to managed identity
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $managedIdentity.Id -BodyParameter $bodyParam

    }

    else {

        Write-Output "Permission '$permission' already exists for managed identity '$($managedIdentity.DisplayName)'"

    }

}


