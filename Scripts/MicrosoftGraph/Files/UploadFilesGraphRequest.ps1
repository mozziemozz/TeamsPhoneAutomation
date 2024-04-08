$siteName = "Azure Automation"

Connect-MgGraph

# Site
$site = Get-MgSite -Search $siteName
$siteId = $site.id

# Drives
$drives = Get-MgSiteDrive -SiteId $siteId -All
$drive = $drives | Where-Object { $_.Name -eq "Documents" -or $_.Name -eq "Dokumente" }

# SharePoint Drive Id (Document Library)
$driveId = $drive.id

# File to upload
$filePath = "C:\Temp\Test.txt"
$fileProperties = Get-ChildItem -Path $filePath

# Read the file content as a byte array
$fileContent = [System.IO.File]::ReadAllBytes($filePath)

# Destination file name
$destinationName = "$($fileProperties.BaseName)-$(Get-Date -Format "yyyy-MM-dd HH-mm-ss").$($fileProperties.Extension)"

# Upload the file to SharePoint
Invoke-GraphRequest -Method PUT -Uri "https://graph.microsoft.com/v1.0/drives/$driveId/root:/Test/$destinationName`:/content" -Body $fileContent -ContentType "application/octet-stream" -Headers $Header
