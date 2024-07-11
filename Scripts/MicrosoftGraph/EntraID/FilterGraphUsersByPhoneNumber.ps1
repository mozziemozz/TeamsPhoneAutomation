<#
    .SYNOPSIS
    Example requests to filter Entra ID users by phone number.

    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.0
    Changelog:
                        2024-07-11: Initial release
    
    Website:            https://heusser.pro
    Sponsor Project:    https://github.com/sponsors/mozziemozz
    Buy me Coffee:      https://buymeacoffee.com/martin.heusser

    This script inlcudes an example for all supported filter queries to get users by business or mobile phone number

    .EXAMPLE
    .\FilterGraphUsersByPhoneNumber.ps1

#>

Connect-MgGraph

# Example numbers
$mobileNumber = "+41 79 456 78 90"
$businessNumber = "+41 43 123 45 67"

$mobileNumberUrlSafe = $mobileNumber.Replace("+", "%2B")
$businessNumberUrlSafe = $businessNumber.Replace("+", "%2B")

$header = @{
    ConsistencyLevel = "eventual"
}

# Filter users by mobile phone number exact match (equals)
$mgUserMatchesMobilePhone = (Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users?`$filter=mobilePhone eq '$mobileNumberUrlSafe'&`$count=true" -ContentType "application/json" -Headers $header).value

# Filter users by business phone number exact match (equals)
$mgUserMatchesBusinessPhone = (Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users?`$filter=businessPhones/any(p:p eq '$businessNumberUrlSafe')&`$count=true" -ContentType "application/json" -Headers $header).value

# Filter users by mobile phone number not equals
$mgUserMatchesMobilePhone = (Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users?`$filter=mobilePhone ne '$mobileNumberUrlSafe'&`$count=true" -ContentType "application/json" -Headers $header).value

# Filter users by business phone number not equals
$mgUserMatchesBusinessPhone = (Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users?`$filter=not businessPhones/any(p:p eq '$businessNumberUrlSafe')&`$count=true" -ContentType "application/json" -Headers $header).value

# Filter users by mobile phone number starts with
$mgUserMatchesMobilePhone = (Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users?`$filter=startsWith(mobilePhone, '$mobileNumberUrlSafe')&`$count=true" -ContentType "application/json" -Headers $header).value  

# Filter users by business phone number starts with
$mgUserMatchesBusinessPhone = (Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users?`$filter=businessPhones/any(p:startswith(p,'$businessNumberUrlSafe'))&`$count=true" -ContentType "application/json" -Headers $header).value

# Filter users by mobile phone number greater or equal
$mgUserMatchesMobilePhone = (Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users?`$filter=mobilePhone ge '$mobileNumberUrlSafe'&`$count=true" -ContentType "application/json" -Headers $header).value

# Filter users by business phone number greater or equal
$mgUserMatchesBusinessPhone = (Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users?`$filter=businessPhones/any(p:p ge '$businessNumberUrlSafe')&`$count=true" -ContentType "application/json" -Headers $header).value

# Filter users by mobile phone number less or equal
$mgUserMatchesMobilePhone = (Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users?`$filter=mobilePhone le '$mobileNumberUrlSafe'&`$count=true" -ContentType "application/json" -Headers $header).value

# Filter users by business phone number less or equal
$mgUserMatchesBusinessPhone = (Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users?`$filter=businessPhones/any(p:p le '$businessNumberUrlSafe')&`$count=true" -ContentType "application/json" -Headers $header).value

# Filter users by mobile phone number in list
$mobileNumberUrlSafe2 = "+41 79 456 78 91".Replace("+", "%2B")
$mgUserMatchesMobilePhone = (Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users?`$filter=mobilePhone in ['$mobileNumberUrlSafe','$mobileNumberUrlSafe2']&`$count=true" -ContentType "application/json" -Headers $header).value

# Convert the results to PS custom objects
$mgUserMatchesMobilePhone = $mgUserMatchesMobilePhone | ConvertTo-Json | ConvertFrom-Json
$mgUserMatchesBusinessPhone = $mgUserMatchesBusinessPhone | ConvertTo-Json | ConvertFrom-Json