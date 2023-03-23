# Re-uploads a new version of the Direct Routing number source list to Azure. Existing numbers are overwritten. The new list must contain all Direct Routing numbers.

$newEnvironment = Get-Content .\Resources\Environment.json | ConvertFrom-Json

$tenantId = $newEnvironment.TenantId
$automationAccountName = $newEnvironment.AutomationAccountName
$resourceGroupName = $newEnvironment.ResourceGroupName

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

# Get local direct routing numbers
$localDirectRoutingNumbers = "'" + (Import-Csv -Path .\Resources\DirectRoutingNumbers.csv | ConvertTo-Json | Out-String) + "'"

# Get remote direct routing numbers from automation variable

$remoteDirectRoutingNumbers = (Get-AzAutomationVariable -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -Name "TeamsPhoneNumberOverview_DirectRoutingNumbers").Value

if ($localDirectRoutingNumbers -ne $remoteDirectRoutingNumbers) {

    Write-Host "Local and remote Direct Routing Numbers do not match. Uploading the new numbers to Azure now..." -ForegroundColor Magenta
    Set-AzAutomationVariable -AutomationAccountName $automationAccountName -Name "TeamsPhoneNumberOverview_DirectRoutingNumbers" -Encrypted $false -Value $localDirectRoutingNumbers -ResourceGroupName $resourceGroupName

}

else {

    Write-Host "Local and remote Direct Routing Numbers match. No need to update." -ForegroundColor Cyan

}

