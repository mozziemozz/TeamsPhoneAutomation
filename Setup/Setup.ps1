$banner = @"
 ______   ______     ______     __    __     ______        __   __     __  __     __    __     ______     ______     ______        __         __     ______     ______  
/\__  _\ /\  ___\   /\  __ \   /\ "-./  \   /\  ___\      /\ "-.\ \   /\ \/\ \   /\ "-./  \   /\  == \   /\  ___\   /\  == \      /\ \       /\ \   /\  ___\   /\__  _\ 
\/_/\ \/ \ \  __\   \ \  __ \  \ \ \-./\ \  \ \___  \     \ \ \-.  \  \ \ \_\ \  \ \ \-./\ \  \ \  __<   \ \  __\   \ \  __<      \ \ \____  \ \ \  \ \___  \  \/_/\ \/ 
   \ \_\  \ \_____\  \ \_\ \_\  \ \_\ \ \_\  \/\_____\     \ \_\\"\_\  \ \_____\  \ \_\ \ \_\  \ \_____\  \ \_____\  \ \_\ \_\     \ \_____\  \ \_\  \/\_____\    \ \_\ 
    \/_/   \/_____/   \/_/\/_/   \/_/  \/_/   \/_____/      \/_/ \/_/   \/_____/   \/_/  \/_/   \/_____/   \/_____/   \/_/ /_/      \/_____/   \/_/   \/_____/     \/_/ 
                                                                                                                                                                        
"@

Write-Host $banner -ForegroundColor DarkCyan

Write-Host "Press enter to start the deployment." -ForegroundColor Cyan
Read-Host

if (!(Test-Path -Path .\.local)) {

    New-Item -Path .\.local -ItemType Directory

}

$repoPath = Get-Content -Path C:\Temp\RepoPath.txt

Set-Location -Path $repoPath

# Import functions
. .\Functions\Connect-MgGraphHTTP.ps1

$newEnvironment = Get-Content .\Resources\Environment.json | ConvertFrom-Json

$tenantId = $newEnvironment.TenantId
$automationAccountName = $newEnvironment.AutomationAccountName
$resourceGroupName = $newEnvironment.ResourceGroupName
$azLocation = $newEnvironment.AzLocation
$groupId = $newEnvironment.GroupId
$MsListName = $newEnvironment.MSListName

$scheduledRunbookTags = @{"Service"="Microsoft Teams";"RunbookType"="Scheduled"}
$functionRunbookTags = @{"Service"="Azure Automation";"RunbookType"="Function"}
$resourceGroupTags = @{"Service"="Azure Automation"}

$installedModules = Get-InstalledModule

$requiredModules = @(

    "Az.Accounts",
    "Az.Automation",
    "Az.Resources",
    "MicrosoftTeams"

)

$missingModules = @()

foreach ($module in $requiredModules) {

    if ($installedModules.Name -contains $module) {

        Write-Host "$module is already installed." -ForegroundColor Green

        if ($module -eq "MicrosoftTeams") {

            $localTeamsPSVersion = Get-InstalledModule -Name "MicrosoftTeams"
            $localTeamsPSVersionMajor = $localTeamsPSVersion.Version.Major
            $localTeamsPSVersionMinor =$localTeamsPSVersion.Version.Minor

            if ($localTeamsPSVersionMajor -lt 5 -and $localTeamsPSVersionMinor -lt 1) {

                Write-Warning -Message "The installed version of $module is older than 5.1.0. $module will be updated."

                function Test-Admin {
                    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
                    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
                }    
            
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

                Update-Module -Name "MicrosoftTeams" -Force

            }

            else {
                
                Write-Host "$module version is new enough for Service Principal Authentication." -ForegroundColor Green

            }

        }

    }

    else {

        Write-Host "$module is not installed." -ForegroundColor Yellow

        $missingModules += $module

    }

}

if ($missingModules) {

    function Test-Admin {
        $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
        $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    }    

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

    Write-Host "Finished installing required modules." -ForegroundColor Cyan
    
}

# Check if Node.JS is installed

try {

    $ErrorActionPreference = "SilentlyContinue"
    $checkNPM = npm list -g
    $ErrorActionPreference = "Continue"

    if ($checkNPM) {

        Write-Host "Node.JS is already installed." -ForegroundColor Green

    }

}
catch {

    # Install Node.JS

    Write-Host "Node.JS is not installed." -ForegroundColor Yellow
    Write-Host "Attempting to install Node.JS..." -ForegroundColor Cyan

    winget install --id=OpenJS.NodeJS  -e

    if ($?) {

        Write-Host "Finished installing Node.JS." -ForegroundColor Cyan

        # Reload path environment variable
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine")

    }

    else {

        Write-Error -Message "Error installing Node.JS. Please install it manually."

    }

}

$checkNPM = npm list -g

if ($checkNPM -match "@pnp/cli-microsoft365") {

    Write-Host "CLI for Microsoft 365 is already installed." -ForegroundColor Green

}

else {

    Write-Host "CLI for Microsoft 365 is not installed." -ForegroundColor Yellow
    Write-Host "Attempting to install CLI for Microsoft 365..." -ForegroundColor Cyan

    npm i -g @pnp/cli-microsoft365

    if ($?) {

        Write-Host "Finished installing CLI for Microsoft 365." -ForegroundColor Cyan

        # Reload path environment variable
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine")

    }

    else {

        Write-Error -Message "Error installing CLI for Microsoft 365. Please install it manually."

    }

}


# Login to npx m365 CLI
npx m365 login

$AADAppRegistrationName = "$automationAccountName" + "_AAD_SP"

$newAADApp = npx m365 aad app add --name $AADAppRegistrationName --withSecret

$encryptedClientSecret = ConvertTo-SecureString ($newAADApp | ConvertFrom-Json).secrets.value -AsPlainText -Force

Set-Content -Path .\.local\AppSecret.txt -Value ($encryptedClientSecret | ConvertFrom-SecureString).TrimEnd() -Force

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

# $newAppObjectId = (($newAADApp | ConvertFrom-Json).objectId).TrimEnd()
$newAppId = (($newAADApp | ConvertFrom-Json).appId).TrimEnd()

Set-Content -Path .\.local\AppId.txt -Value $newAppId.TrimEnd() -Force

Write-Host "Adding the required Microsoft Graph permissions to $AADAppRegistrationName..." -ForegroundColor Cyan

$newAADServicePrincipalObjectId = ((npx m365 aad sp add --appId $newAppId | ConvertFrom-Json).Id).TrimEnd()

npx m365 aad approleassignment add --resource "Microsoft Graph" --scope "Organization.Read.All" --appId $newAppId
npx m365 aad approleassignment add --resource "Microsoft Graph" --scope "User.Read.All" --appId $newAppId
npx m365 aad approleassignment add --resource "Microsoft Graph" --scope "Channel.Delete.All" --appId $newAppId
npx m365 aad approleassignment add --resource "Microsoft Graph" --scope "ChannelSettings.ReadWrite.All" --appId $newAppId
npx m365 aad approleassignment add --resource "Microsoft Graph" --scope "Group.ReadWrite.All" --appId $newAppId
npx m365 aad approleassignment add --resource "Microsoft Graph" --scope "ChannelMember.ReadWrite.All" --appId $newAppId
npx m365 aad approleassignment add --resource "Microsoft Graph" --scope "AppCatalog.ReadWrite.All" --appId $newAppId
npx m365 aad approleassignment add --resource "Microsoft Graph" --scope "TeamSettings.ReadWrite.All" --appId $newAppId
npx m365 aad approleassignment add --resource "Microsoft Graph" --scope "Sites.Manage.All" --appId $newAppId
npx m365 aad approleassignment add --resource "Microsoft Graph" --scope "RoleManagement.ReadWrite.Directory" --appId $newAppId

$AppId = Get-Content -Path .\.local\AppId.txt | Out-String
$AppSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR((Get-Content -Path .\.local\AppSecret.txt | ConvertTo-SecureString))) | Out-String

do {
    
    Write-Host "Sleeping for 30s to give Azure some time to apply the permissions..." -ForegroundColor Cyan

    Start-Sleep 30

    . Connect-MgGraphHTTP -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret

    $directoryRoleId = ((Invoke-RestMethod -Method Get -Headers $Header -Uri "https://graph.microsoft.com/v1.0/directoryRoles/").value | Where-Object {$_.displayName -eq "Skype for Business Administrator"}).Id

} until (
    ($?)
)


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

New-AzResourceGroup -Name $resourceGroupName -Location $azLocation -Tag $resourceGroupTags

New-AzAutomationAccount -Name $automationAccountName -Location $azLocation -ResourceGroupName $resourceGroupName -Tags $resourceGroupTags

$installedTeamsPSVersion = (Get-InstalledModule -Name MicrosoftTeams).Version -join ""

# Install MicrosoftTeams PowerShell Module in Automation Account
New-AzAutomationModule -AutomationAccountName $automationAccountName -Name "MicrosoftTeams" -ContentLink "https://psg-prod-eastus.azureedge.net/packages/microsoftteams.$installedTeamsPSVersion.nupkg" -ResourceGroupName $resourceGroupName

do {

    Write-Host "Checking if the module has finished importing. Next check in 1 minute..." -ForegroundColor Cyan

    Start-Sleep 60

    $checkModuleImportProgress = Get-AzAutomationModule -AutomationAccountName $automationAccountName -ResourceGroupName $resourceGroupName -Name "MicrosoftTeams"
    
} until (
    $checkModuleImportProgress.ProvisioningState -eq "Succeeded"
)


# Upload variables

New-AzAutomationVariable -AutomationAccountName $automationAccountName -Name "TeamsPhoneNumberOverview_AppId" -Encrypted $false -Value $AppId -ResourceGroupName $resourceGroupName
New-AzAutomationVariable -AutomationAccountName $automationAccountName -Name "TeamsPhoneNumberOverview_AppSecret" -Encrypted $true -Value $AppSecret -ResourceGroupName $resourceGroupName

New-AzAutomationVariable -AutomationAccountName $automationAccountName -Name "TeamsPhoneNumberOverview_GroupId" -Encrypted $false -Value $groupId -ResourceGroupName $resourceGroupName
New-AzAutomationVariable -AutomationAccountName $automationAccountName -Name "TeamsPhoneNumberOverview_TenantId" -Encrypted $false -Value $tenantId -ResourceGroupName $resourceGroupName

# Upload Country Lookup Table
$countryLookupTable = "'" + (Get-Content -Path .\Resources\CountryLookupTable.json | Out-String) + "'"
New-AzAutomationVariable -AutomationAccountName $automationAccountName -Name "TeamsPhoneNumberOverview_CountryLookupTable" -Encrypted $false -Value $countryLookupTable -ResourceGroupName $resourceGroupName

# Upload Country Lookup Table
$createListJson = "'" + (Get-Content -Path .\Resources\CreateList.json | Out-String) + "'"
New-AzAutomationVariable -AutomationAccountName $automationAccountName -Name "TeamsPhoneNumberOverview_CreateList" -Encrypted $false -Value $createListJson -ResourceGroupName $resourceGroupName

# Upload alldirectroutingnumbers
$directRoutingNumbers = "'" + (Import-Csv -Path .\Resources\DirectRoutingNumbers.csv | ConvertTo-Json | Out-String) + "'"
New-AzAutomationVariable -AutomationAccountName $automationAccountName -Name "TeamsPhoneNumberOverview_DirectRoutingNumbers" -Encrypted $false -Value $directRoutingNumbers -ResourceGroupName $resourceGroupName

# Upload MS List name
New-AzAutomationVariable -AutomationAccountName $automationAccountName -Name "TeamsPhoneNumberOverview_MsListName" -Encrypted $false -Value $MsListName -ResourceGroupName $resourceGroupName

# Upload function runbooks

$functionRunbooks = Get-ChildItem -Path .\Functions | Where-Object {$_.Name -notmatch "Local"}

foreach ($functionRunbook in $functionRunbooks) {

    $newRunbook = Import-AzAutomationRunbook -Path $functionRunbook.FullName -Tags $functionRunbookTags -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -Type PowerShell -Name $($functionRunbook.Name.Replace(".ps1","")) -Published

}

# Upload main runbook

$mainRunbook = Get-ChildItem -Path .\Scripts\TeamsPhoneNumberOverview.ps1

$newRunbook = Import-AzAutomationRunbook -Path $mainRunbook.FullName -Tags $scheduledRunbookTags -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -Type PowerShell -Name $($mainRunbook.Name.Replace(".ps1","")) -Published

$nextFullHour = (Get-Date).Hour
$startTime = (Get-Date "$($nextFullHour):00:00").ToUniversalTime().AddHours(2)

# Upload schedule

$newSchedule = New-AzAutomationSchedule -AutomationAccountName $automationAccountName -Name $functionRunbook.Name.Replace(".ps1","") -StartTime $StartTime -HourInterval 1 -ResourceGroupName $resourceGroupName -TimeZone "Etc/UTC"

Register-AzAutomationScheduledRunbook -RunbookName $newRunbook.Name -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -ScheduleName $newSchedule.Name

Remove-Item -Path C:\Temp\RepoPath.txt

Write-Host "Finished provisioning Azure Automation Account! Press enter to exit." -ForegroundColor Cyan
Read-Host