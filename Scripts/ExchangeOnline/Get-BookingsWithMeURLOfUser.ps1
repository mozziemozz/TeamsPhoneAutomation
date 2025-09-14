<#

    .SYNOPSIS
        Get the Microsoft Bookings URL of a user in Exchange Online.

    .DESCRIPTION
        This script retrieves the Microsoft Bookings URL for a user in Exchange Online. It checks if a user has Microsoft Bookings enabled and if they have set up their personal bookings page already. If a user has not set up their personal bookings page, it will be reflected in the output.

    .NOTES
        Author:     Martin Heusser
        Link:       https://heusser.pro
                    https://github.com/sponsors/mozziemozz?frequency=one-time&sponsor=mozziemozz
                    https://buymeacoffee.com/martin.heusser
        Version:    1.0.0
        Date:       2025-06-21

#>

$exoConnection = Get-ConnectionInformation

if (!$exoConnection) {

    Connect-ExchangeOnline -ShowBanner:$false

}

$tenantId = $exoConnection.TenantID

$graphSession = Get-MgUser -Top 1 -ErrorAction SilentlyContinue

if (!$graphSession) {

    Connect-MgGraph -Scopes "User.Read.All" -NoWelcome

}

if (-not $allUsers) {

    $allUsers = Get-MgUser -Filter "assignedPlans/any(s:s/servicePlanId eq 199a5c09-e0ca-4e37-8f7c-b05d533e1ea2)" -Property Id, UserPrincipalName, DisplayName -CountVariable Count -ConsistencyLevel eventual -All

}

$userPrincipalName = $allUsers | Select-Object DisplayName, Id, UserPrincipalName | Out-GridView -Title "Select a user to get their Microsoft Bookings URL (Note: Only Bookings Licensed users are shown)" -PassThru | Select-Object -ExpandProperty UserPrincipalName

$domain = $userPrincipalName.Split("@")[1]

$exoMailBox = Get-EXOMailbox -Identity $userPrincipalName -Properties ExchangeGUID

$mailBox = Get-Mailbox -Identity $userPrincipalName

if ($mailBox.PersistedCapabilities -notcontains "BPOS_S_BookingsAddOn") {

    Write-Host "User '$($exoMailBox.DisplayName)' does not have Microsoft Bookings enabled in Exchange Online." -ForegroundColor Red

}

else {

    Write-Host "User '$($exoMailBox.DisplayName)' has Microsoft Bookings enabled in Exchange Online." -ForegroundColor Green

    $exchangeGUIDWithHyphens = $exoMailBox.ExchangeGUID.Guid

    $exchangeGUID = $exoMailBox.ExchangeGUID.Guid.Replace("-", "")

    $serviceId = $null
    $response = $null
    $checkPersonalBookingsPage = $null
    $personalBookingsURL = $null
    $checkPersonalBookingsPageBody = $null

    $personalBookingsURL = "https://outlook.office.com/bookwithme/user/$exchangeGUID%40$domain"

    try {

        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $session.Cookies.Add((New-Object System.Net.Cookie("ClientId", "4DB0129438C34891BDD6F2D6141584B9", "/", "outlook.office.com")))
        $session.Cookies.Add((New-Object System.Net.Cookie("OIDC", "1", "/", "outlook.office.com")))
        $response = Invoke-WebRequest -UseBasicParsing -Uri "https://outlook.office.com/BookingsService/api/V1/bookingBusinessesc2/mbx:$($exchangeGUID)@$($tenantId)/services" `
        -WebSession $session `
        -Headers @{
        "authority"="outlook.office.com"
        "method"="GET"
        "path"="/BookingsService/api/V1/bookingBusinessesc2/mbx:$($exchangeGUID)@$($tenantId)/services"
        "scheme"="https"
        "accept"="*/*"
        "accept-encoding"="gzip, deflate, br, zstd"
        "accept-language"="en-US,en;q=0.9"
        "prefer"="exchange.behavior=`"IncludeThirdPartyOnlineMeetingProviders`""
        "priority"="u=1, i"
        "sec-ch-ua-mobile"="?0"
        "sec-ch-ua-platform"="`"Windows`""
        "sec-fetch-dest"="empty"
        "sec-fetch-mode"="cors"
        "sec-fetch-site"="same-origin"
        "x-anchormailbox"="mbx:$($exchangeGUID)@$($tenantId)"
        "x-edge-shopping-flag"="0"
        "x-owa-canary"="X-OWA-CANARY_cookie_is_null_or_empty"
        "x-owa-hosted-ux"="false"
        "x-req-source"="BookWithMe"
        } `
        -ContentType "application/json; charset=utf-8" -ErrorAction Stop

        # $serviceId = $response.Content | ConvertFrom-Json | Select-Object -ExpandProperty service | Where-Object { $_.isPrivate -eq $true } | Select-Object -ExpandProperty serviceId -First 1
        $serviceId = $response.Content | ConvertFrom-Json | Select-Object -ExpandProperty service | Select-Object -ExpandProperty serviceId -First 1

        $checkPersonalBookingsPageBody = [pscustomobject]@{
            staffIds      = @("1c1c7887-e0fc-479b-ac7d-6983d0f026ff") 
            startDateTime = @{
                dateTime = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
                timeZone = (Get-TimeZone).Id
            }
            endDateTime   = @{
                dateTime = (Get-Date).AddDays(1).ToString("yyyy-MM-ddTHH:mm:ss")
                timeZone = (Get-TimeZone).Id
            }
            serviceId     = $serviceId
        } | ConvertTo-Json -Depth 5 -Compress
        
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $session.Cookies.Add((New-Object System.Net.Cookie("ClientId", "09A51AD7CB19435B9DCFE57421FB3DD6", "/", "outlook.office.com")))
        $session.Cookies.Add((New-Object System.Net.Cookie("OIDC", "1", "/", "outlook.office.com")))
        $checkPersonalBookingsPage = Invoke-WebRequest -UseBasicParsing -Uri "https://outlook.office.com/BookingsService/api/V1/bookingBusinessesc2/mbx:$($exchangeGUID)@$($tenantId)/getStaffAvailability" `
        -Method "POST" `
        -WebSession $session `
        -Headers @{
            "authority"="outlook.office.com"
            "method"="POST"
            "path"="/BookingsService/api/V1/bookingBusinessesc2/mbx:$($exchangeGUID)@$($tenantId)/getStaffAvailability"
            "scheme"="https"
            "accept"="*/*"
            "accept-encoding"="gzip, deflate, br, zstd"
            "accept-language"="de-CH,de;q=0.9"
            "cache-control"="no-cache"
            "dnt"="1"
            "origin"="https://outlook.office.com"
            "pragma"="no-cache"
            "prefer"="exchange.behavior=`"IncludeThirdPartyOnlineMeetingProviders`""
            "priority"="u=1, i"
            "sec-ch-ua-mobile"="?0"
            "sec-fetch-dest"="empty"
            "sec-fetch-mode"="cors"
            "sec-fetch-site"="same-origin"
            "x-anchormailbox"="mbx:$($exchangeGUID)@$($tenantId)"
            "x-edge-shopping-flag"="0"
            "x-owa-canary"="X-OWA-CANARY_cookie_is_null_or_empty"
            "x-owa-hosted-ux"="false"
            "x-req-source"="BookWithMe"
        } `
        -ContentType "application/json; charset=utf-8" `
        -Body $checkPersonalBookingsPageBody -ErrorAction Stop

        Write-Host "User '$($exoMailBox.DisplayName)' set up their Personal Bookings page already." -ForegroundColor Green

        $personalBookingsPageConfigured = $true

    }
    catch {
        
        Write-Host "User '$($exoMailBox.DisplayName)' didn't set up their Personal Bookings page yet." -ForegroundColor Yellow

        $personalBookingsPageConfigured = $false

    }


    if ($personalBookingsPageConfigured) {

        $personalBookingsURLAnonymous = "$($personalBookingsURL)?anonymous"

        $personalBookingsURLAnonymous | Set-Clipboard

        Write-Host "Microsoft Bookings URL of user '$($exoMailBox.DisplayName)' is '$personalBookingsURL'." -ForegroundColor Cyan

        Write-Host "Microsoft Bookings anonymous URL of user '$($exoMailBox.DisplayName)': is '$personalBookingsURLAnonymous'" -ForegroundColor Magenta

        Write-Host "The URL has been copied to the clipboard." -ForegroundColor Green

    }

}