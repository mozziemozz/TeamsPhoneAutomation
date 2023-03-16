# From: https://adamtheautomator.com/powershell-graph-api/

function Connect-MgGraphHTTP {
    param (
        [Parameter(Mandatory=$true)][String]$TenantId,
        [Parameter(Mandatory=$true)][String]$AppId,
        [Parameter(Mandatory=$true)][String]$AppSecret
    )

    # Define AppId, secret and scope, your tenant name and endpoint URL
    $Scope = "https://graph.microsoft.com/.default"
    $Url = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    # Add System.Web for urlencode
    Add-Type -AssemblyName System.Web
    # Create body
    $Body = @{
        client_id = $AppId
        client_secret = $AppSecret
        scope = $Scope
        grant_type = 'client_credentials'
    }
    # Splat the parameters for Invoke-Restmethod for cleaner code
    $PostSplat = @{
        ContentType = 'application/x-www-form-urlencoded'
        Method = 'POST'
        # Create string by joining bodylist with '&'
        Body = $Body
        Uri = $Url
    }
    # Request the token!
    $Request = Invoke-RestMethod @PostSplat
    # Create header
    $Header = @{
        Authorization = "$($Request.token_type) $($Request.access_token)"
    }

}
