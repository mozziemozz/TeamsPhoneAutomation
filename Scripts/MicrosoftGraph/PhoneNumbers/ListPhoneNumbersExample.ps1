Connect-MgGraph -Scopes "TeamsUserConfiguration.Read.All", "TeamsTelephoneNumber.ReadWrite.All"

# Get all user configurations in Microsoft Teams using the Graph API
$allPages = @()

$userConfigurations = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta//admin/teams/userConfigurations" -ContentType "application/json"
$allPages += $userConfigurations.value

if ($userConfigurations.'@odata.nextLink') {

    do {

        $userConfigurations = (Invoke-MgGraphRequest -Method Get -Uri $userConfigurations.'@odata.nextLink' -ContentType "application/json")
        $allPages += $userConfigurations.value

    } until (
        !$userConfigurations.'@odata.nextLink'
    )
        
}

$userConfigurations = ($allPages | ConvertTo-Json -Depth 99 | ConvertFrom-Json -Depth 99)

# Filter example
$phoneNumber = ""
$filterByPhoneNumber = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/admin/teams/userConfigurations?`$filter=telephoneNumbers/any(p:p/telephoneNumber eq '$($phoneNumber)')&`$count=true" -ContentType "application/json"

# Private Preview
# Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/admin/teams/telephoneNumberManagement/NumberAssignments" -ContentType "application/json"