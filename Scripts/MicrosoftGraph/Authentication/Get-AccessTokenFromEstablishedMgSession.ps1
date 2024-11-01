# Connect to Microsoft Graph using interactive browser sign-in
Connect-MgGraph

# Example user id
$userId = ""

# Get the user's profile photo content via Microsoft Graph PowerShell SDK
# This will only allow you to download the profile photo to disk but not store it in memory
Get-MgUserPhotoContent -UserId $userId -OutFile "C:\Temp\profilePhoto.jpg"

# Make a request using Invoke-MgGraphRequest using the session established by Connect-MgGraph
$mgRequest = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$($userId)/photo/`$value" -ContentType "image/jpeg" -OutputType HttpResponseMessage

# Create a header with the access token from the session for subsequent requests via Invoke-RestMethod or Invoke-WebRequest
$authHeader = @{
    Authorization = "$($mgRequest.RequestMessage.Headers.Authorization.Scheme) $($mgRequest.RequestMessage.Headers.Authorization.Parameter)"
}

# Make a request using Invoke-WebRequest with the access token from the session
# This will allow you to download the profile photo to memory and store it in a variable as a byte array
$profilePhoto = (Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/users/$($userId)/photo/`$value" -Headers $authHeader).Content

$alternativeRequest = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me" -OutputType HttpResponseMessage).RequestMessage.Headers.Authorization.Parameter