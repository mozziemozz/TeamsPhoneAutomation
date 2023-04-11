function Get-MZZLicenseState {
    param (
       [Parameter(Mandatory=$true)][string]$UserPrincipalName
    )
 
       (Get-MgUserLicenseDetail -UserId $UserPrincipalName).ServicePlans
    
}

function Sync-MZZCQAgents {
    param (
        [Parameter(Mandatory=$false)][string]$CQIdentity
    )

    . Connect-MZZTeams

    if ($CQIdentity) {

        $selectedCQs = Get-CsCallQueue -WarningAction SilentlyContinue -Identity $CQIdentity

    }

    else {

        # 100 is the default
        $CQs = Get-CsCallQueue -WarningAction SilentlyContinue -First 100
    
        if ($CQs.Count -ge 100) {

            Write-Host "This tenant has more than 100 Call Queues. Querrying additional CQs..." -ForegroundColor Cyan
    
            $skipCounter = 100
    
            do {
        
                $querriedCQs = Get-CsCallQueue -WarningAction SilentlyContinue -Skip $skipCounter
        
                $CQs += $querriedCQs
    
                $skipCounter += $querriedCQs.Count
    
            } until (
                $querriedCQs.Count -eq 0
            )
    
        }

        $selectedCQs = $CQs | Select-Object Name, Identity | Out-GridView -Title "Choose one or multiple Call Queues from the list..." -PassThru

    }

    foreach ($CQ in $selectedCQs) {

        $existingCQAgents = . Get-MZZCQAgents -CQIdentity $CQ.Identity

        Set-CsCallQueue -Identity $CQ.Identity -WarningAction SilentlyContinue > $null

        $updatedCQAgents = . Get-MZZCQAgents -CQIdentity $CQ.Identity

        $removedMembers = $existingCQAgents | Where-Object {$updatedCQAgents."User Principal Name" -notcontains $_."User Principal Name"}
        $addedMembers = $updatedCQAgents | Where-Object {$existingCQAgents."User Principal Name" -notcontains $_."User Principal Name"}

        Write-Host "The Agent List of Call Queue '$($CQ.Name)' has been force synced." -ForegroundColor Cyan

        if ($removedMembers) {

            Write-Host "The following Agents were removed:" -ForegroundColor Magenta
            $removedMembers | Out-String

        }

        if ($addedMembers) {

            Write-Host "The following Agents were added:" -ForegroundColor Magenta
            $addedMembers | Out-String

        }

        if (!$removedMembers -and !$addedMembers) {

            Write-Host "No new or remove Agents were detected. Please try again in a few minutes." -ForegroundColor Magenta

        }

    }
    
}

function Get-MZZCQAgents {
    param (
        [Parameter(Mandatory=$false)][string]$CQIdentity
    )

    . Connect-MZZTeams

    if ($CQIdentity) {

        $cqAgents = (Get-CsCallQueue -WarningAction SilentlyContinue -Identity $CQIdentity).Agents

    }

    else {

        # 100 is the default
        $CQs = Get-CsCallQueue -WarningAction SilentlyContinue -First 100

        if ($CQs.Count -ge 100) {

            Write-Host "This tenant has more than 100 Call Queues. Querrying additional CQs..." -ForegroundColor Cyan
    
            $skipCounter = 100
    
            do {
        
                $querriedCQs = Get-CsCallQueue -WarningAction SilentlyContinue -Skip $skipCounter
        
                $CQs += $querriedCQs
    
                $skipCounter += $querriedCQs.Count
    
            } until (
                $querriedCQs.Count -eq 0
            )
    
        }

        do {

            $selectedCQ = $CQs | Select-Object Name, Identity | Out-GridView -Title "Choose a Call Queue from the list..." -PassThru

            if ($selectedCQ.Name.Count -gt 1) {

                Write-Warning -Message "You have selected more than 1 Call Queue. You can only select one. Please try again."

            }
            
        } until (
            $selectedCQ.Name.Count -eq 1
        )

		Write-Host "Selected Call Queue: $($selectedCQ.Name)" -ForegroundColor Cyan
        
        $cqAgents = (Get-CsCallQueue -WarningAction SilentlyContinue -Identity $selectedCQ.Identity).Agents
        
    }
 
 
    $agentList = @()
 
    foreach ($agent in $cqAgents) {
        
        $agentProperties = New-Object -TypeName psobject
 
        $teamsAgent = Get-CsOnlineUser -Identity $agent.ObjectId
 
        $agentProperties | Add-Member -MemberType NoteProperty -Name "User Principal Name" -Value $teamsAgent.UserPrincipalName
        $agentProperties | Add-Member -MemberType NoteProperty -Name "OptIn Status" -Value $agent.OptIn
 
        $agentList += $agentProperties
 
    }

    return $agentList
  
}

function Connect-MZZTeams {
    param (
    )

    try {
        $msTeamsTenant = Get-CsTenant -ErrorAction Stop > $null
        $msTeamsTenant = Get-CsTenant
    }
    catch {
        Connect-MicrosoftTeams -ErrorAction SilentlyContinue > $null
        $msTeamsTenant = Get-CsTenant
    }
    finally {
        if ($msTeamsTenant -and $? -eq $true) {

            Write-Host "Connected Teams Tenant: $($msTeamsTenant.DisplayName)" -ForegroundColor Green
            
        }

        if ($msTeamsTenant.TenantId -is [System.ValueType]) {

            $msTeamsTenantId = $msTeamsTenant.TenantId.Guid

        }

        else {

            $msTeamsTenantId = $msTeamsTenant.TenantId

        }
    }
    
}