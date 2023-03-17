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

        $missingModules += $module

    }

}

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

    Import-AzAutomationRunbook -Path $functionRunbook.FullName -Tags $functionRunbookTags -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -Type PowerShell -Name $($functionRunbook.Name.Replace(".ps1","")) -Published

}
