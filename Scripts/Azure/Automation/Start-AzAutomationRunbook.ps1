Connect-AzAccount

# Define environment
$resourceGroupName = "mzz-rmg-013"
$automationAccountName = "mzz-automation-account-014"
$runbookName = "New-TeamsResourceAccount"

# Define the parameters for the runbook

$params = @{
    UserPrincipalName = "ra_aa_az-automation-001@nocaptech.ch"
    DisplayName = "Az Automation 001"
    VoiceAppType = "AutoAttendant"
    UsageLocation = "CH"   
}

# Start the runbook
$startJob = Start-AzAutomationRunbook -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -Name $runbookName -Parameters $params

# Get the job status

$counter = 0

do {

    Start-Sleep -Seconds 30

    $job = Get-AzAutomationJob -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -JobId $startJob.JobId

    $counter += 30
} until (
    $job.Status -eq "Completed" -or $counter -ge 300
)

switch ($job.Status) {
    Completed {

        Write-Host "Job completed successfully!"

    }
    Stopped {

        Write-Host "Job failed to complete!"

    }
    Failed {

        Write-Host "Job failed to complete!"

    }
    Default {}
}