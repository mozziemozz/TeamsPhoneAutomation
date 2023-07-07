# Version: 2.1

# Set to true if script is executed locally
$localTestMode = $false

function Get-AllSPOListItems {
    param (
        [Parameter(Mandatory=$true)][string]$ListId
    )
    
        # Get existing list items
        $sharePointListItems = @()

        $querriedItems = (Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists/$($ListId)/items?expand=fields")
        $sharePointListItems += $querriedItems.value.fields
    
        if ($querriedItems.'@odata.nextLink') {
    
            Write-Output "List contains more than $($querriedItems.value.Count) items. Querrying additional items..."
    
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
        $MsListName = "Teams Phone Number Overview Demo V2"
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

# $userInformationListId = (Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists?`$filter=displayName eq '$localizedUserInformationList'").value.id

# From: https://stackoverflow.com/questions/61143146/how-to-get-user-from-user-field-lookupid
$userInformationListId = ((Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists?select=id,name,system").value | Where-Object {$_.name -eq "users"}).id

# Retrieve all list items
. Get-AllSPOListItems -ListId $userInformationListId
$userInformationList = $sharePointListItems

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

. Get-AllSPOListItems -ListId $sharePointListId

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

                $assignReservedNumber = $false
    
            }

            else {

                $assignReservedNumber = $true
    
            }    

        }

        else {

            if ($userPrincipalName) {

                $assignReservedNumber = $true

            }

            else {

                Write-Output "User with lookup id $($reservedNumber.UserProfileLookupId) is not available in the lookup table yet. The number will be assigned in the next job."

            }

        }

        if ($assignReservedNumber -eq $true) {

            Write-Output "Checking License and Usage Location for user $userPrincipalName..."

            switch ($reservedNumber.Number_x0020_Type) {
                CallingPlan {

                    # Check if user has Calling Plan license, no need for Teams Phone Standard Check because CP requires Teams Phone Standard
                    if ($checkCsOnlineUser.FeatureTypes -contains "CallingPlan") {

                        $licenseCheckSuccess = $true

                    }

                    else {

                        $licenseCheckSuccess = $false

                    }

                    if ($checkCsOnlineUser.UsageLocation -eq $reservedNumber.Country) {

                        $usageLocationCheck = $true

                    }

                    else {

                        $usageLocationCheck = $false

                    }

                    $assignVoiceRoutingPolicy = $false

                }

                OperatorConnect {

                    # Check if user has Teams Phone Standard License
                    if ($checkCsOnlineUser.FeatureTypes -contains "PhoneSystem") {

                        $licenseCheckSuccess = $true

                        if ($checkCsOnlineUser.UsageLocation -eq $reservedNumber.Country) {

                            $usageLocationCheck = $true

                        }

                        else {

                            $usageLocationCheck = $false

                        }


                    }

                    else {

                        $licenseCheckSuccess = $false

                    }

                    $assignVoiceRoutingPolicy = $false

                }

                DirectRouting {

                    # Check if user has Teams Phone Standard License
                    if ($checkCsOnlineUser.FeatureTypes -contains "PhoneSystem") {

                        $licenseCheckSuccess = $true

                        $usageLocationCheck = $true

                    }

                    else {

                        $licenseCheckSuccess = $false

                    }

                    $assignVoiceRoutingPolicy = $true

                }
                Default {}
            }

            # Trying to fix usage location errors
            if ($usageLocationCheck -eq $false) {

                # Usage location does not match phone number
                $patchBody = @{UsageLocation=$($reservedNumber.Country)} | ConvertTo-Json

                Invoke-RestMethod -Method Patch -Headers $Header -Uri "https://graph.microsoft.com/v1.0/users/$($checkCsOnlineUser.Identity)" -ContentType "application/json" -Body $patchBody
                
                if ($?) {

                    do {
                        Write-Output "Sleeping for 20s..."

                        Start-Sleep 20

                        $checkCsOnlineUser = Get-CsOnlineUser -Identity $checkCsOnlineUser.Identity
                    } until (
                        $checkCsOnlineUser.UsageLocation -eq $reservedNumber.Country
                    )

                    Write-Output "Usage location has been successfully changed to $($checkCsOnlineUser.UsageLocation) for user $($checkCsOnlineUser.UserPrincipalName)"

                    $usageLocationCheck = $true
                }

                else {

                    Write-Output "Error while trying to change usage location to $($checkCsOnlineUser.UsageLocation) for user $($checkCsOnlineUser.UserPrincipalName)"

                    $usageLocationCheck = $false

                }

            }

            if ($licenseCheckSuccess -eq $true -and $usageLocationCheck -eq $true) {

                Write-Output "License and Usage Location checks for user $userPrincipalName are successful."
                Write-Output "Trying to assign reserved number $($reservedNumber.Title) to user $userPrincipalName..."

                Set-CsPhoneNumberAssignment -Identity $userPrincipalName -PhoneNumberType $reservedNumber.Number_x0020_Type -PhoneNumber $reservedNumber.Title

                if ($assignVoiceRoutingPolicy -eq $true) {

                    $phoneNumber = $reservedNumber.Title

                    . Get-CountryFromPrefix

                    Write-Output "$($reservedNumber.Title) is a Direct Routing Number. Voice Routing Policy $voiceRoutingPolicy will be assigned."

                    Grant-CsOnlineVoiceRoutingPolicy -Identity $userPrincipalName -PolicyName $voiceRoutingPolicy

                }

            }

            # License issue
            else {

                if ($licenseCheckSuccess -eq $false) {

                    Write-Output "User $userPrincipalName is missing the license for $($reservedNumber.Number_x0020_Type) assignment."

                }

                if ($usageLocationCheck -eq $false) {

                    Write-Output "Usage Location of $userPrincipalName is $($checkCsOnlineUser.UsageLocation) and does not match phone number country $($reservedNumber.Country)."

                }

                ($sharePointListItems | Where-Object {$_.Title -eq $reservedNumber.Title}).Status = "Assignment Error"

            }

        }

    }
    
}

# Add leading plus ("+") to all numbers
$allDirectRoutingNumbers = $allDirectRoutingNumbers | ForEach-Object {"+" + $_.PhoneNumber}

# Get CsOnline Numbers
$allCsOnlineNumbers = . Get-CsOnlineNumbers

# Get all Teams users which have a phone number assigned
# $allTeamsPhoneUsers = Get-CsOnlineUser -Filter "LineURI -ne `$null" -ResultSize 9999
$allTeamsPhoneUsers = Get-CsOnlineUser -Filter "(FeatureTypes -contains 'PhoneSystem') -or (FeatureTypes -contains 'VoiceApp')" -ResultSize 9999 | Select-Object Identity, UserPrincipalName, DisplayName, LineURI, FeatureTypes
$allTeamsPhoneUserDetails = @()

$userCounter = 1

foreach ($teamsPhoneUser in $allTeamsPhoneUsers) {

    if ($userCounter % 100 -eq 0) {

        Write-Output "Working on $userCounter/$($allTeamsPhoneUsers.Count)..."

    }

    # Check if user has a LineURI

    if (!$teamsPhoneUser.LineUri) {

        $teamsPhoneUser = Get-CsOnlineUser -Identity $teamsPhoneUser.Identity | Select-Object Identity, UserPrincipalName, DisplayName, LineURI, FeatureTypes

    }

    if ($teamsPhoneUser.LineUri) {

        $teamsPhoneUserDetails = New-Object -TypeName psobject

        if ($teamsPhoneUser.FeatureTypes -contains "VoiceApp") {

            $teamsPhoneUserType = "Resource Account"

        }

        else {

            $teamsPhoneUserType = "User Account"

        }

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

            $assignedDirectRoutingNumberCity = ($allCsOnlineNumbers | Where-Object {$_.TelephoneNumber -eq $phoneNumber}).City

            $numberType = "DirectRouting"
            $city = $assignedDirectRoutingNumberCity

            $country = . Get-CountryFromPrefix

        }

        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Status" -Value "Assigned"
        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Number_x0020_Type" -Value $numberType
        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "City" -Value $city
        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Country" -Value $country

        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "User_x0020_Name" -Value $teamsPhoneUser.DisplayName
        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "User_x0020_Principal_x0020_Name" -Value $teamsPhoneUser.UserPrincipalName
        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "Account_x0020_Type" -Value $teamsPhoneUserType
        $teamsPhoneUserDetails | Add-Member -MemberType NoteProperty -Name "UserId" -Value $teamsPhoneUser.Identity

        $userCounter ++

        $allTeamsPhoneUserDetails += $teamsPhoneUserDetails

    }

}

# Get all unassigned Calling Plan and Operator Connect phone numbers or all conference assigned numbers
foreach ($csOnlineNumber in $allCsOnlineNumbers | Where-Object {$_.PstnAssignmentStatus -eq "ConferenceAssigned" -or ($null -eq $_.AssignedPstnTargetId -and $_.NumberType -ne "DirectRouting")}) {

    $csOnlineNumberDetails = New-Object -TypeName psobject

    $phoneNumber = $csOnlineNumber.TelephoneNumber

    if ($csOnlineNumber.PstnAssignmentStatus -eq "ConferenceAssigned") {

        $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Title" -Value $phoneNumber
        $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Phone_x0020_Extension" -Value "N/A"
        $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Status" -Value "Assigned"
        $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Number_x0020_Type" -Value $csOnlineNumber.NumberType
        $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "City" -Value $csOnlineNumber.City
        $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Country" -Value $csOnlineNumber.IsoCountryCode
    
    
        $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "User_x0020_Name" -Value "Conference Bridge"
        $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "User_x0020_Principal_x0020_Name" "Conference Bridge"

        $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Account_x0020_Type" -Value "Conference Bridge"
        $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "UserId" -Value "Conference Bridge"

    }

    else {

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
    
        $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "Account_x0020_Type" -Value $accountType
        $csOnlineNumberDetails | Add-Member -MemberType NoteProperty -Name "UserId" -Value "Unassigned"   

    }

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

    if ($updateCounter % 100 -eq 0) {

        Write-Output "Working on $updateCounter/$($allTeamsPhoneUserDetails.Count)..."

    }

    if ($sharePointListItems.Title -contains $teamsPhoneNumber."Title") {

        $checkEntryIndex = $sharePointListItems.Title.IndexOf($teamsPhoneNumber.Title)
        $checkEntry = $sharePointListItems[$checkEntryIndex]

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

            # Write-Output "Entry $($teamsPhoneNumber.Title) is already up to date and won't be updated..."

        }

        else {

            if ($checkEntry.Status -eq "Reserved" -and $teamsPhoneNumber.Status -eq "Unassigned") {

                Write-Output "Entry $($teamsPhoneNumber.Title) is reserved and will not be updated..."

            }

            else {

                if ($checkEntry.Status -eq "Assignment Error") {

                    $teamsPhoneNumber = $checkEntryObject

                    Write-Output "Entry $($teamsPhoneNumber.Title) is NOT up to date because it has assignment errors. Entry won't be updated..."


                }

                else {

                    Write-Output "Entry $($teamsPhoneNumber.Title) is NOT up to date and will be updated..."

                }


                # patch

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

        # Only create list item if title is not empty
        if ($teamsPhoneNumber.Title) {

            Invoke-RestMethod -Method Post -Headers $header -ContentType "application/json; charset=UTF-8" -Body $body -Uri "https://graph.microsoft.com/v1.0/sites/$($sharePointSite.id)/lists/$($sharePointListId)/items"

        }
        
    }

    $updateCounter ++

    # Read-Host

}