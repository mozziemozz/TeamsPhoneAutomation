$TenantId = ""
$AppId = ""
$AppSecret = ""
$SiteName = "Azure Automation"

. .\Functions\Connect-MgGraphHTTP.ps1

. Connect-MgGraphHTTP -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret

# Site
$site = Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/v1.0/sites?search=$siteName"
$siteId = $site.value.id

# Drives
$drives = Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives"

# SharePoint Drive Id (Document Library)
$driveId = ($drives.value | Where-Object { $_.Name -eq "Documents" -or $_.Name -eq "Dokumente" }).Id

# File to upload
$filePath = "C:\Temp\Test.txt"
$fileProperties = Get-ChildItem -Path $filePath

# Read the file content as a byte array
$fileContent = [System.IO.File]::ReadAllBytes($filePath)

# Destination file name
$destinationName = "$($fileProperties.BaseName)-$(Get-Date -Format "yyyy-MM-dd HH-mm-ss").$($fileProperties.Extension)"

# Upload the file to SharePoint
Invoke-RestMethod -Method PUT -Uri "https://graph.microsoft.com/v1.0/drives/$driveId/root:/Test/$destinationName`:/content" -Body $fileContent -ContentType "application/octet-stream" -Headers $Header
