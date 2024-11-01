Connect-ExchangeOnline

$userId = Read-Host "Enter UPN"

$domain = $userId.Split("@")[1]

$exoMailBox = Get-EXOMailbox -Identity $userId -Properties ExchangeGUID

$mailBox = Get-Mailbox -Identity $userId

if ($mailBox.PersistedCapabilities -notcontains "BPOS_S_BookingsAddOn") {

    Write-Host "User '$($exoMailBox.DisplayName)' does not have a Bookings with Me license." -ForegroundColor Red

}

else {

    Write-Host "User '$($exoMailBox.DisplayName)' has a Bookings with Me license." -ForegroundColor Green

    $exchangeGUID = $exoMailBox.ExchangeGUID.Guid.Replace("-", "")

    $bookingsWithMeURL = "https://outlook.office.com/bookwithme/user/$exchangeGUID%40$domain"

    $bookingsWithMeURL | Set-Clipboard

    Write-Host "Bookings with me URL of user '$($exoMailBox.DisplayName)' is '$bookingsWithMeURL'. Note: URL copied to clipboard." -ForegroundColor Green

    Write-Host "Tip: For anonymous access, use: '$($bookingsWithMeURL)?anonymous'" -ForegroundColor Yellow

    Start-Process $bookingsWithMeURL

}