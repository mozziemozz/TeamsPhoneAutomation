Set-Location C:\Temp\GitHub\TeamsPhoneAutomation

. .\Functions\Test-Admin.ps1

$TenantId = Get-Content -Path .\.local\TenantId.txt
$automationAccountName = Get-Content -Path .\.local\AutomationAccountName.txt
$resourceGroupName = Get-Content -Path .\.local\ResourceGroupName.txt

$scheduledRunbookTags = @{"Service"="Microsoft Teams";"RunbookType"="Scheduled"}
$functionRunbookTags = @{"Service"="Azure Automation";"RunbookType"="Function"}

$installedModules = Get-InstalledModule

$requiredModules = @(

    "Az.Automation"

)

$missingModules = @()

foreach ($module in $requiredModules) {

    if ($installedModules.Name -contains $module) {

        Write-Host "$module is already installed." -ForegroundColor Green

    }

    else {

        Write-Host "$module is not installed." -ForegroundColor Yellow

        $missingModules += $module

    }

}

if ($missingModules) {

    if ((Test-Admin) -eq $false)  {
        if ($elevated) 
        {
            # tried to elevate, did not work, aborting
        } 
        else {
            Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    }
    
        exit
    }
    
    foreach ($module in $missingModules) {

        Write-Host "Attempting to install $module..." -ForegroundColor Cyan

        Install-Module -Name $module

    }

    Write-Host "Finished installing required modules. Restatarting script now..." -ForegroundColor Cyan

    . $myinvocation.MyCommand.Definition
    
}

# Check if Node.JS is installed

try {

    $ErrorActionPreference = "SilentlyContinue"
    $checkNPM = npm list -g
    $ErrorActionPreference = "Continue"

    if ($checkNPM) {

        Write-Host "Node.JS is already installed." -ForegroundColor Green

        if ($checkNPM -match "@pnp/cli-microsoft365") {

            Write-Host "CLI for Microsoft 365 is already installed." -ForegroundColor Green

        }

        else {

            Write-Host "CLI for Microsoft 365 is not installed." -ForegroundColor Yellow
            Write-Host "Attempting to install CLI for Microsoft 365..." -ForegroundColor Cyan

            npm i -g @pnp/cli-microsoft365

            if ($?) {

                Write-Host "Finished installing CLI for Microsoft 365. Restatarting script now..." -ForegroundColor Cyan

                . $myinvocation.MyCommand.Definition        

            }

            else {

                Write-Error -Message "Error installing CLI for Microsoft 365. Please install it manually."

            }

        }

    }

}
catch {

    # Install Node.JS

    Write-Host "Node.JS is not installed." -ForegroundColor Yellow
    Write-Host "Attempting to install Node.JS..." -ForegroundColor Cyan

    winget install --id=OpenJS.NodeJS  -e

    if ($?) {

        Write-Host "Finished installing Node.JS. Restatarting script now..." -ForegroundColor Cyan

        # Reload path environment variable
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine")

        . $myinvocation.MyCommand.Definition

    }

    else {

        Write-Error -Message "Error installing Node.JS. Please install it manually."

    }

}

# Login to M365 CLI
m365 login

$AADAppRegistrationName = "$automationAccountName" + "_AAD_SP"

$newAADApp = m365 aad app add --name $AADAppRegistrationName --withSecret

$encryptedClientSecret = ConvertTo-SecureString ($newAADApp | ConvertFrom-Json).secrets.value -AsPlainText -Force

Set-Content -Path .\.local\AppSecret.txt -Value ($encryptedClientSecret | ConvertFrom-SecureString) -Force

Write-Host "Client secret saved to \.local\AppSecret.txt in an encrypted state." -ForegroundColor Cyan

$reviewClientSecret = Read-Host "Would you like to display the Client secret? [Y] / [N] ?"

switch ($reviewClientSecret) {
    y {

        Write-Host "Here's your Client secret. It has been decrypted using this machine and user $($env:USERNAME). The client secret can only be decrypted on this machine with this account!" -ForegroundColor Cyan
        
        $importedClientSecret = Get-Content -Path .\.local\AppSecret.txt | ConvertTo-SecureString

        [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($importedClientSecret))

        Write-Host "Script will continue in 15s. Please copy the secret and store it in a safe place." -ForegroundColor Yellow
        Start-Sleep 15

    }
    n {
        
        Write-Host "Client secret won't be displayed." -ForegroundColor Magenta

    }
    Default {}
}

$newAppObjectId = (($newAADApp | ConvertFrom-Json).objectId).Trim()
$newAppId = (($newAADApp | ConvertFrom-Json).appId).Trim()

Write-Host "Adding the required Microsoft Graph permissions to $AADAppRegistrationName..." -ForegroundColor Cyan

$newAADServicePrincipalObjectId = (m365 aad sp add --appId $newAppId | ConvertFrom-Json).Id

m365 aad approleassignment add --resource "Microsoft Graph" --scope "Organization.Read.All" --appId $newAppId
m365 aad approleassignment add --resource "Microsoft Graph" --scope "User.Read.All" --appId $newAppId
m365 aad approleassignment add --resource "Microsoft Graph" --scope "Channel.Delete.All" --appId $newAppId
m365 aad approleassignment add --resource "Microsoft Graph" --scope "ChannelSettings.ReadWrite.All" --appId $newAppId
m365 aad approleassignment add --resource "Microsoft Graph" --scope "Group.ReadWrite.All" --appId $newAppId
m365 aad approleassignment add --resource "Microsoft Graph" --scope "ChannelMember.ReadWrite.All" --appId $newAppId
m365 aad approleassignment add --resource "Microsoft Graph" --scope "AppCatalog.ReadWrite.All" --appId $newAppId
m365 aad approleassignment add --resource "Microsoft Graph" --scope "TeamSettings.ReadWrite.All" --appId $newAppId
m365 aad approleassignment add --resource "Microsoft Graph" --scope "Sites.Manage.All" --appId $newAppId
m365 aad approleassignment add --resource "Microsoft Graph" --scope "RoleManagement.ReadWrite.Directory" --appId $newAppId

. .\Functions\Connect-MgGraphHTTP.ps1

$TenantId = Get-Content -Path .\.local\TenantId.txt
$AppId = Get-Content -Path .\.local\AppId.txt
$AppSecret = Get-Content -Path .\.local\AppSecret.txt

. Connect-MgGraphHTTP -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret

$directoryRoleId = ((Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/v1.0/directoryRoles/").value | Where-Object {$_.displayName -eq "Skype for Business Administrator"}).Id

$addDirectoryRoleBody = @"
{
    "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/$newAADServicePrincipalObjectId"
}
"@

Invoke-RestMethod -Method Post -Headers $Header -ContentType "application/json" -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$directoryRoleId/members/`$ref" -Body $addDirectoryRoleBody

# Connect to Azure Account
$checkAzureSession = Get-AzContext > $null

if (!$checkAzureSession.TenantId -eq $tenantId) {

    Connect-AzAccount -TenantId $tenantId

    $azSubscriptions = Get-AzSubscription

    if ($azSubscriptions.Count -gt 1) {

        Write-Host "Multiple Azure Subcriptions found. Please choose one from the list..." -ForegroundColor Cyan

        $azSubscription = $azSubscriptions | Out-GridView -Title "Choose a Subscription From The List" -PassThru

        Select-AzSubscription $azSubscription.Id

    }

}

# Upload runbooks

$functionRunbooks = Get-ChildItem -Path .\Functions

foreach ($functionRunbook in $functionRunbooks) {

    $newRunbook = Import-AzAutomationRunbook -Path $functionRunbook.FullName -Tags $functionRunbookTags -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -Type PowerShell -Name $($functionRunbook.Name.Replace(".ps1","")) -Published

    if ($functionRunbook.Name -eq "Get-CsOnlineNumbers.ps1") {

        $nextFullHour = (Get-Date).Hour
        $startTime = (Get-Date "$($nextFullHour):00:00").ToUniversalTime().AddHours(2)

        $newSchedule = New-AzAutomationSchedule -AutomationAccountName $automationAccountName -Name $functionRunbook.Name.Replace(".ps1","") -StartTime $StartTime -HourInterval 1 -ResourceGroupName $resourceGroupName -TimeZone "Etc/UTC"

        Register-AzAutomationScheduledRunbook -RunbookName $newRunbook.Name -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -ScheduleName $newSchedule.Name

    }

}
