# Import external functions
. .\Functions\Connect-MsTeamsServicePrincipal.ps1
. .\Functions\Connect-MgGraphHTTP.ps1
. .\Functions\Get-CountryFromPrefix.ps1
. .\Functions\Get-CsOnlineNumbers.ps1

$MsListName = "Teams Phone Number Overview 3"

$TenantId = Get-Content -Path .\.local\TenantId.txt
$AppId = Get-Content -Path .\.local\AppId.txt
$AppSecret = Get-Content -Path .\.local\AppSecret.txt

$groupId = Get-Content -Path .\.local\GroupId.txt

. Connect-MsTeamsServicePrincipal -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret

. Connect-MgGraphHTTP -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret

# Import Direct Routing numbers
$allDirectRoutingNumbers = Import-Csv -Path .\Resources\DirectRoutingNumbers.csv -Encoding UTF8

# Add leading plus ("+") to all numbers
$allDirectRoutingNumbers = $allDirectRoutingNumbers | ForEach-Object {"+" + $_.PhoneNumber}

# Get CsOnline Numbers
$allCsOnlineNumbers = . Get-CsOnlineNumbers

# Get existing SharePoint lists for group id
$sharePointSite = (Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/sites/root")

$existingSharePointLists = (Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists").value

if ($existingSharePointLists.name -contains $MsListName) {

    Write-Output "A list with the name $MsListName already exists in site $($sharePointSite.name). No new list will be created."

    $sharePointListId = ($existingSharePointLists | Where-Object {$_.Name -eq $MsListName}).id

}

else {

    Write-Output "A list with the name $MsListName does not exist in site $($sharePointSite.name). A new list will be created."

    $createListJson = (Get-Content -Path .\Resources\CreateList.json).Replace("Name Placeholder",$MsListName)

    $newSharePointList = Invoke-RestMethod -Method Post -Headers $Header -ContentType "application/json" -Body $createListJson -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists"

    $sharePointListId = $newSharePointList.id

}

# Get all Teams users which have a phone number assigned
$allTeamsPhoneUsers = Get-CsOnlineUser -Filter "LineURI -ne `$null"

$allTeamsPhoneUserDetails = @()

foreach ($teamsPhoneUser in $allTeamsPhoneUsers) {

    $teamsPhoneUserDetails = New-Object -TypeName psobject

    if ($teamsPhoneUser.FeatureTypes -contains "VoiceApp") {

        $teamsPhoneUserType = "Resource Account"

    }

    else {

        $teamsPhoneUserType = "User Account"

    }

    if ($teamsPhoneUser.LineUri) {

        $phoneNumber = $teamsPhoneUser.LineUri.Replace("tel:","")

        $country = . Get-CountryFromPrefix

        if ($teamsPhoneUser.LineUri -match ";") {

            $lineUri = $teamsPhoneUser.LineUri.Split(";")[0]
            $extension = $teamsPhoneUser.LineUri.Split(";")[-1]

            $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Title" -Value $lineUri.Replace("tel:","")
            $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Phone_x0020_Extension" -Value $extension

        }

        else {

            $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Title" -Value $teamsPhoneUser.LineUri.Replace("tel:","")
            $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Phone_x0020_Extension" -Value "N/A"
        }

        if ($allCsOnlineNumbers.PhoneNumber -contains $phoneNumber) {

            $numberType = ($allCsOnlineNumbers | Where-Object {$_.TelephoneNumber -eq ($teamsPhoneUser.LineUri).Replace("tel:","")}).NumberType

        }

        else {

            $numberType = "DirectRouting"

        }
        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Status" -Value "Assigned"
        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Number_x0020_Type" -Value $numberType
        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Country" -Value $country


    }

    $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "User_x0020_Name" -Value $teamsPhoneUser.DisplayName
    $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "User_x0020_Principal_x0020_Name" -Value $teamsPhoneUser.UserPrincipalName
    $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Account_x0020_Type" -Value $teamsPhoneUserType
    $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "UserId" -Value $teamsPhoneUser.Identity
    # $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "TeamsAdminCenter" -Value "https://admin.teams.microsoft.com/users/$($teamsPhoneUser.Identity)/account"

    $allTeamsPhoneUserDetails += $teamsPhoneUserDetails

}

# Get all unassigned Calling Plan and Operator Connect phone numbers
foreach ($csOnlineNumber in $allCsOnlineNumbers | Where-Object {$null -eq $_.AssignedPstnTargetId -and $_.NumberType -ne "DirectRouting"}) {

    $csOnlineNumberDetails = New-Object -TypeName psobject

    $phoneNumber = $csOnlineNumber.TelephoneNumber

    $country = . Get-CountryFromPrefix

    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Title" -Value $phoneNumber
    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Phone_x0020_Extension" -Value "N/A"
    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Status" -Value "Unassigned"
    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Number_x0020_Type" -Value $csOnlineNumber.NumberType
    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Country" -Value $country


    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "User_x0020_Name" -Value "Unassigned"
    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "User_x0020_Principal_x0020_Name" "Unassigned"

    if ($csOnlineNumber.Capability -contains "UserAssignment") {

        $accountType = "User Account"

    }

    elseif ($csOnlineNumber.Capability -contains "VoiceApplicationAssignment" -and $csOnlineNumber.Capability -notcontains "ConferenceAssignment") {

        $accountType = "Resource Account"

    }

    elseif ($csOnlineNumber.Capability -notcontains "VoiceApplicationAssignment" -and $csOnlineNumber.Capability -contains "ConferenceAssignment") {

        $accountType = "Conference Bridge"

    }

    else {

        $accountType = "Resource Account, Conference Bridge"

    }

    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Account_x0020_Type" -Value  $accountType
    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "UserId" -Value "Unassigned"
    # $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "TeamsAdminCenter" -Value $null


    $allTeamsPhoneUserDetails += $csOnlineNumberDetails

}

# Get all unassigned Direct Routing Numbers
$directRoutingNumbers = $allDirectRoutingNumbers | Where-Object {$allTeamsPhoneUserDetails."Title" -notcontains $_ }

foreach ($directRoutingNumber in $directRoutingNumbers) {

    $directRoutingNumberDetails = New-Object -TypeName psobject

    $phoneNumber = $directRoutingNumber

    $country = . Get-CountryFromPrefix

    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "Title" -Value $directRoutingNumber
    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "Phone_x0020_Extension" -Value "N/A"
    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "Status" -Value "Unassigned"
    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "Number_x0020_Type" -Value "DirectRouting"
    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "Country" -Value $country


    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "User_x0020_Name" -Value "Unassigned"
    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "User_x0020_Principal_x0020_Name" "Unassigned"
    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "Account_x0020_Type" -Value "User Account, Resource Account"
    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "UserId" -Value "Unassigned"
    # $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "TeamsAdminCenter" -Value $null

    $allTeamsPhoneUserDetails += $directRoutingNumberDetails

}

# Get existing list items
$sharePointListItems = (Invoke-RestMethod -Method Get -Headers $header -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists/$($sharePointListId)/items?expand=fields").value.fields

# Update list
foreach ($teamsPhoneNumber in $allTeamsPhoneUserDetails) {

    if ($sharePointListItems.Title -contains $teamsPhoneNumber."Title" -and $teamsPhoneNumber."Title" -ne "Unassigned") {

        # entry already existis in list checking if data is up to date

        $itemId = ($sharePointListItems | Where-Object {$_.Title -eq $teamsPhoneNumber.Title}).id

        $checkEntry = (Invoke-RestMethod -Method Get -Headers $header -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists/$($sharePointListId)/items/$itemId`?expand=fields")

        $checkEntryObject = New-Object -TypeName psobject

        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "Title" -Value $checkEntry.fields.Title
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "Phone_x0020_Extension" -Value $checkEntry.fields.Phone_x0020_Extension
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "Status" -Value $checkEntry.fields.Status
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "Number_x0020_Type" -Value $checkEntry.fields.Number_x0020_Type
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "Country" -Value $checkEntry.fields.Country
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "User_x0020_Name" -Value $checkEntry.fields.User_x0020_Name
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "User_x0020_Principal_x0020_Name" -Value $checkEntry.fields.User_x0020_Principal_x0020_Name
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "Account_x0020_Type" -Value $checkEntry.fields.Account_x0020_Type
        # $checkEntryObject | Add-Member -MemberType NoteProperty -Name "TeamsAdminCenter" -Value $checkEntry.fields.TeamsAdminCenter


        $compareObjects = ($checkEntryObject | Out-String) -eq ($teamsPhoneNumber | Out-String)

        if ($compareObjects) {

            # no differences

            Write-Host "Entry $($teamsPhoneNumber.Title) is up to date..."

        }

        else {

            # patch

            Write-Host "Entry $($teamsPhoneNumber.Title) is NOT up to date..."

$body = @"
{
"fields": 

"@
            
                    $body += ($teamsPhoneNumber | ConvertTo-Json)
                    $body += "`n}"
            
                    Invoke-RestMethod -Method Patch -Headers $header -ContentType "application/json" -Body $body -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists/$($sharePointListId)/items/$itemId"
            

        }

    }

    else {

        # entry does not exist in list


        $body = @"
{
"fields": 

"@

        $body += ($teamsPhoneNumber | ConvertTo-Json)
        $body += "`n}"

        Invoke-RestMethod -Method Post -Headers $header -ContentType "application/json" -Body $body -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists/$($sharePointListId)/items"

    }

}


or(and(not(equals(triggerOutputs()?['body/User_x0020_Principal_x0020_Name'], 'Unassigned')),not(contains(triggerBody(), 'UserProfile')), equals(triggerOutputs()?['body/Account_x0020_Type'], 'User Account')),not(equals(triggerOutputs()?['body/User_x0020_Principal_x0020_Name'],triggerOutputs()?['body/UserProfile']['Email'])))