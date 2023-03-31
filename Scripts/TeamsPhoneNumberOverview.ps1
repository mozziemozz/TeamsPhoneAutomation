$localTestMode = $true

function Get-AllSPOListItems {
    param (
        
    )
    
        # Get existing list items
        $sharePointListItems = @()

        $querriedItems = (Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists/$($sharePointListId)/items?expand=fields")
        $sharePointListItems += $querriedItems.value.fields
    
        if ($querriedItems.'@odata.nextLink') {
    
            Write-Output "List contains more than $($querriedItems.value.Count) itesm. Querrying additional items..."
    
            do {
    
                $querriedItems = (Invoke-RestMethod -Method Get -Headers $Header -Uri $querriedItems.'@odata.nextLink')
                $sharePointListItems += $querriedItems.value.fields
                
            } until (
                !$querriedItems.'@odata.nextLink'
            )
    
        }
    
        else {
    
            Write-Output "All items were retrieved in the first request."
    
        }
    
        Write-Output "Finished retrieving $($sharePointListItems.Count) items."
    
}

switch ($localTestMode) {
    $true {

        # Local Environment

        # Import external functions
        . .\Functions\Connect-MsTeamsServicePrincipal.ps1
        . .\Functions\Connect-MgGraphHTTP.ps1
        . .\Functions\Get-CountryFromPrefix.ps1
        . .\Functions\Get-CsOnlineNumbers.ps1

        # Import variables
        $MsListName = "Teams Phone Number Demo 10"
        $TenantId = Get-Content -Path .\.local\TenantId.txt
        $AppId = Get-Content -Path .\.local\AppId.txt
        $AppSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR((Get-Content -Path .\.local\AppSecret.txt | ConvertTo-SecureString))) | Out-String
        $groupId = Get-Content -Path .\.local\GroupId.txt

        # Import Direct Routing numbers
        $allDirectRoutingNumbers = Import-Csv -Path .\Resources\DirectRoutingNumbers.csv -Encoding UTF8
        
    }

    $false {

        # Azure Automation

        # Import external functions
        . .\Connect-MsTeamsServicePrincipal.ps1
        . .\Connect-MgGraphHTTP.ps1
        . .\Get-CountryFromPrefix.ps1
        . .\Get-CsOnlineNumbers.ps1

        # Import variables        
        $MsListName =  Get-AutomationVariable -Name "TeamsPhoneNumberOverview_MsListName"
        $TenantId = Get-AutomationVariable -Name "TeamsPhoneNumberOverview_TenantId"
        $AppId = Get-AutomationVariable -Name "TeamsPhoneNumberOverview_AppId"
        $AppSecret = Get-AutomationVariable -Name "TeamsPhoneNumberOverview_AppSecret"
        $groupId = Get-AutomationVariable -Name "TeamsPhoneNumberOverview_GroupId"

        # Import Direct Routing numbers
        $allDirectRoutingNumbers = (Get-AutomationVariable -Name "TeamsPhoneNumberOverview_DirectRoutingNumbers").Replace("'","") | ConvertFrom-Json

    }
    Default {}
}

. Connect-MsTeamsServicePrincipal -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret

. Connect-MgGraphHTTP -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret

# Get existing SharePoint lists for group id
$sharePointSite = (Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/sites/root")

$existingSharePointLists = (Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists").value


# https://mozzism.sharepoint.com/sites/AzureAutomation/_catalogs/users/simple.aspx

$userInformationListId = (Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists?`$filter=displayName eq 'Benutzerinformationsliste'").value.id

$userInformationList = (Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists/$userInformationListId/items?expand=fields").value.fields

$userLookupIds = $userInformationList | Select-Object Username,UserSelection

if ($existingSharePointLists.name -contains $MsListName) {

    Write-Output "A list with the name $MsListName already exists in site $($sharePointSite.name). No new list will be created."

    $sharePointListId = ($existingSharePointLists | Where-Object {$_.Name -eq $MsListName}).id

}

else {

    Write-Output "A list with the name $MsListName does not exist in site $($sharePointSite.name). A new list will be created."

    switch ($localTestMode) {
        $true {
    
            # Local Environment
    
            $createListJson = (Get-Content -Path .\Resources\CreateList.json).Replace("Name Placeholder",$MsListName)
            
        }
    
        $false {
    
            # Azure Automation
    
            $createListJson = (Get-AutomationVariable -Name "TeamsPhoneNumberOverview_CreateList").Replace("Name Placeholder",$MsListName).Replace("'","")

        }
        Default {}
    }
    
    $newSharePointList = Invoke-RestMethod -Method Post -Headers $Header -ContentType "application/json" -Body $createListJson -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists"

    $sharePointListId = $newSharePointList.id

}

. Get-AllSPOListItems

if ($sharePointListItems) {

    # Unassign numbers

    foreach ($reservedNumber in ($sharePointListItems | Where-Object {$_.Status -eq "Remove Pending" -and $_.User_x0020_Principal_x0020_Name -ne "Unassigned"})) {

        Write-Output "Trying to remove the number $($reservedNumber.Title) from user $($reservedNumber.User_x0020_Principal_x0020_Name)..."

        Remove-CsPhoneNumberAssignment -Identity $reservedNumber.User_x0020_Principal_x0020_Name -RemoveAll

    }

    # Assign reserved numbers

    foreach ($reservedNumber in ($sharePointListItems | Where-Object {$_.Status -eq "Reserved" -and $_.UserProfileLookupId -ne $null})) {

        $userPrincipalName = ($userLookupIds | Where-Object {$_.UserSelection -eq $reservedNumber.UserProfileLookupId}).Username

        $checkCsOnlineUser = (Get-CsOnlineUser -Identity $userPrincipalName)

        if ($checkCsOnlineUser.LineURI) {

            $checkCsOnlineUserLineURI = $checkCsOnlineUser.LineURI.Replace("tel:","")

            if ($checkCsOnlineUserLineURI -ne $reservedNumber.Title) {

                Write-Output "User $userPrincipalName already has $checkCsOnlineUserLineURI assigned. Number will be removed and replaced with $($reservedNumber.Title)"
    
                Remove-CsPhoneNumberAssignment -Identity $userPrincipalName -RemoveAll
    
            }
    
            if ($checkCsOnlineUserLineURI -eq $reservedNumber.Title) {
    
                Write-Output "Reserved number $($reservedNumber.Title) is already assigned to $userPrincipalName."
    
            }

            else {

                Write-Output "Trying to assign reserved number $($reservedNumber.Title) to user $userPrincipalName..."
    
                Set-CsPhoneNumberAssignment -Identity $userPrincipalName -PhoneNumberType $reservedNumber.Number_x0020_Type -PhoneNumber $reservedNumber.Title
    
            }    

        }

        else {

            Write-Output "Trying to assign reserved number $($reservedNumber.Title) to user $userPrincipalName..."

            Set-CsPhoneNumberAssignment -Identity $userPrincipalName -PhoneNumberType $reservedNumber.Number_x0020_Type -PhoneNumber $reservedNumber.Title

        }

    }
    
}

# Add leading plus ("+") to all numbers
$allDirectRoutingNumbers = $allDirectRoutingNumbers | ForEach-Object {"+" + $_.PhoneNumber}

# Get CsOnline Numbers
$allCsOnlineNumbers = . Get-CsOnlineNumbers

# Get all Teams users which have a phone number assigned
$allTeamsPhoneUsers = Get-CsOnlineUser -Filter "LineURI -ne `$null"

$allTeamsPhoneUserDetails = @()

$userCounter = 1

foreach ($teamsPhoneUser in $allTeamsPhoneUsers) {

    Write-Output "Working on $userCounter/$($allTeamsPhoneUsers.Count)..."
    
    $teamsPhoneUserDetails = New-Object -TypeName psobject

    if ($teamsPhoneUser.FeatureTypes -contains "VoiceApp") {

        $teamsPhoneUserType = "Resource Account"

    }

    else {

        $teamsPhoneUserType = "User Account"

    }

    if ($teamsPhoneUser.LineUri) {

        $phoneNumber = $teamsPhoneUser.LineUri.Replace("tel:","")

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

        if ($allCsOnlineNumbers.TelephoneNumber -contains $phoneNumber) {

            $matchingCsOnlineNumber = ($allCsOnlineNumbers | Where-Object {$_.TelephoneNumber -eq ($teamsPhoneUser.LineUri).Replace("tel:","")})

            $numberType = $matchingCsOnlineNumber.NumberType
            $city = $matchingCsOnlineNumber.City

            if ($matchingCsOnlineNumber.IsoCountryCode) {
                
                $country = $matchingCsOnlineNumber.IsoCountryCode

            }

            else {

                $country = . Get-CountryFromPrefix

            }

        }

        else {

            $assignedDirectRoutingNumberCity = (Get-CsPhoneNumberAssignment -TelephoneNumber $($teamsPhoneUser.LineUri).Replace("tel:","")).City

            $numberType = "DirectRouting"
            $city = $assignedDirectRoutingNumberCity

            $country = . Get-CountryFromPrefix

        }


        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Status" -Value "Assigned"
        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Number_x0020_Type" -Value $numberType
        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "City" -Value $city
        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Country" -Value $country


    }

    $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "User_x0020_Name" -Value $teamsPhoneUser.DisplayName
    $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "User_x0020_Principal_x0020_Name" -Value $teamsPhoneUser.UserPrincipalName
    $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Account_x0020_Type" -Value $teamsPhoneUserType
    $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "UserId" -Value $teamsPhoneUser.Identity

    $userCounter ++

    $allTeamsPhoneUserDetails += $teamsPhoneUserDetails

}

# Get all unassigned Calling Plan and Operator Connect phone numbers
foreach ($csOnlineNumber in $allCsOnlineNumbers | Where-Object {$null -eq $_.AssignedPstnTargetId -and $_.NumberType -ne "DirectRouting"}) {

    $csOnlineNumberDetails = New-Object -TypeName psobject

    $phoneNumber = $csOnlineNumber.TelephoneNumber

    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Title" -Value $phoneNumber
    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Phone_x0020_Extension" -Value "N/A"
    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Status" -Value "Unassigned"
    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Number_x0020_Type" -Value $csOnlineNumber.NumberType
    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "City" -Value $csOnlineNumber.City
    $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Country" -Value $csOnlineNumber.IsoCountryCode


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
    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "City" -Value "N/A"
    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "Country" -Value $country

    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "User_x0020_Name" -Value "Unassigned"
    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "User_x0020_Principal_x0020_Name" "Unassigned"
    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "Account_x0020_Type" -Value "User Account, Resource Account"
    $directRoutingNumberDetails | Add-Member -MemberType NoteProperty -Name "UserId" -Value "Unassigned"

    $allTeamsPhoneUserDetails += $directRoutingNumberDetails

}

if ($sharePointListItems) {

    foreach ($spoPhoneNumber in $sharePointListItems) {
    
        if ($spoPhoneNumber.Title -notin $allTeamsPhoneUserDetails.Title) {

            Write-Output "Entry $($spoPhoneNumber.Title) is no longer available. It will be removed from the list..."

            Invoke-RestMethod -Method Delete -Headers $Header -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists/$($sharePointListId)/items/$($spoPhoneNumber.id)"

        }

    }

}

# Update list

$updateCounter = 1

foreach ($teamsPhoneNumber in $allTeamsPhoneUserDetails) {

    Write-Output "Working on $updateCounter/$($allTeamsPhoneUserDetails.Count)..."

    # if ($sharePointListItems.Title -contains $teamsPhoneNumber."Title" -and $teamsPhoneNumber."Title" -ne "Unassigned") {

    if ($sharePointListItems.Title -contains $teamsPhoneNumber."Title") {

        # entry already existis in list checking if data is up to date

        # $itemId = ($sharePointListItems | Where-Object {$_.Title -eq $teamsPhoneNumber.Title}).id

        # $checkEntry = (Invoke-RestMethod -Method Get -Headers $header -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists/$($sharePointListId)/items/$itemId`?expand=fields")

        $checkEntry = ($sharePointListItems | Where-Object {$_.Title -eq $teamsPhoneNumber.Title})

        $itemId = $checkEntry.id

        $checkEntryObject = New-Object -TypeName psobject

        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "Title" -Value $checkEntry.Title
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "Phone_x0020_Extension" -Value $checkEntry.Phone_x0020_Extension
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "Status" -Value $checkEntry.Status
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "Number_x0020_Type" -Value $checkEntry.Number_x0020_Type
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "City" -Value $checkEntry.City
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "Country" -Value $checkEntry.Country
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "User_x0020_Name" -Value $checkEntry.User_x0020_Name
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "User_x0020_Principal_x0020_Name" -Value $checkEntry.User_x0020_Principal_x0020_Name
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "Account_x0020_Type" -Value $checkEntry.Account_x0020_Type
        $checkEntryObject | Add-Member -MemberType NoteProperty -Name "UserId" -Value $checkEntry.UserId

        $compareObjects = ($checkEntryObject | Out-String) -eq ($teamsPhoneNumber | Out-String)

        if ($compareObjects) {

            # no differences

            Write-Output "Entry $($teamsPhoneNumber.Title) is up to date..."

        }

        else {

            if ($checkEntry.Status -eq "Reserved" -and $teamsPhoneNumber.Status -eq "Unassigned") {

                Write-Output "Entry $($teamsPhoneNumber.Title) is reserved and will not be updated..."

            }

            else {

                # patch

                Write-Output "Entry $($teamsPhoneNumber.Title) is NOT up to date..."

$body = @"
{
"fields": 

"@
        
                $body += ($teamsPhoneNumber | ConvertTo-Json)
                $body += "`n}"
        
                Invoke-RestMethod -Method Patch -Headers $header -ContentType "application/json; charset=UTF-8" -Body $body -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists/$($sharePointListId)/items/$itemId"
            
            }

        }

    }

    else {

        # entry does not exist in list

        Write-Output "Entry $($teamsPhoneNumber.Title) is NEW..."

        $body = @"
{
"fields": 

"@

        $body += ($teamsPhoneNumber | ConvertTo-Json)
        $body += "`n}"

        Invoke-RestMethod -Method Post -Headers $header -ContentType "application/json; charset=UTF-8" -Body $body -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists/$($sharePointListId)/items"

    }

    $updateCounter ++

    # Read-Host

}