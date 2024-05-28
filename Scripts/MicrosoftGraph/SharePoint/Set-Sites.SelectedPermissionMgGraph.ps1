# Import Graph sites module
Import-Module Microsoft.Graph.Sites

# Connect to Graph
Connect-MgGraph

# Get Site Id by search string
$siteId = (Get-MgSite -Search "Azure Automation").Id

# Get Site Permissions
$sitePermissions = Get-MgSitePermission -SiteId $siteId -Property *

$entraIdAppRegistration = Get-MgApplication -All | Out-GridView -OutputMode Single -Title "Choose an Application Registration from the list"

$params = @{
    roles               = @(
        "write"
    )
    grantedToIdentities = @(
        @{
            application = @{
                id          = "$($entraIdAppRegistration.Id)"
                displayName = "$($entraIdAppRegistration.DisplayName)"
            }
        }
    )
}

New-MgSitePermission -SiteId $siteId -BodyParameter $params
