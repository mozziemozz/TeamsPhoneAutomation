Set-Location C:\Temp\GitHub\TeamsPhoneAutomation

$TenantId = Get-Content -Path .\.local\TenantId.txt
$automationAccountName = Get-Content -Path .\.local\AutomationAccountName.txt
$resourceGroupName = Get-Content -Path .\.local\ResourceGroupName.txt

$scheduledRunbookTags = @{"Service"="Microsoft Teams";"RunbookType"="Scheduled"}
$functionRunbookTags = @{"Service"="Azure Automation";"RunbookType"="Function"}

$installedModules = Get-InstalledModule

$requiredModules = @(

    "Az.Automation",
    "Az.Network"

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

    Write-Host "Finished installing required modules. Restatarting script now..." -ForegroundColor Cyan

    . $myinvocation.MyCommand.Definition
    
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

    $newRunbook = Import-AzAutomationRunbook -Path $functionRunbook.FullName -Tags $functionRunbookTags -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -Type PowerShell -Name $($functionRunbook.Name.Replace(".ps1","")) -Published

    if ($functionRunbook.Name -eq "Get-CsOnlineNumbers.ps1") {

        $nextFullHour = (Get-Date).Hour
        $startTime = (Get-Date "$($nextFullHour):00:00").ToUniversalTime().AddHours(2)

        $newSchedule = New-AzAutomationSchedule -AutomationAccountName $automationAccountName -Name $functionRunbook.Name.Replace(".ps1","") -StartTime $StartTime -HourInterval 1 -ResourceGroupName $resourceGroupName -TimeZone "Etc/UTC"

        Register-AzAutomationScheduledRunbook -RunbookName $newRunbook.Name -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -ScheduleName $newSchedule.Name

    }

}
