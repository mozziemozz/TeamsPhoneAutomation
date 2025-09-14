Write-Warning -Message "This script only works if your auto attendants and call queues have a resource account associated with them!"
Write-Warning -Message "Learn more: https://learn.microsoft.com/en-us/microsoftteams/plan-auto-attendant-call-queue#nested-auto-attendants-and-call-queues"

$reportingDelay = -4

$resourceAccountId = "" # Resource account associated with the top-level auto attendant or call queue

$auth = Get-Content -Path ".\Scripts\Teams\CallRecords\auth.json" | ConvertFrom-Json

$clientSecret = $auth.secret | ConvertTo-SecureString

$clientSecretCredentials = New-Object System.Management.Automation.PSCredential($auth.appId, $clientSecret)

Connect-MgGraph -TenantId $auth.tenantId -Credential $clientSecretCredentials -NoWelcome

Get-MgContext | Format-List AppName, ClientId, TenantId, AuthType, TokenCredentialType, Scopes

# $timeFilter = (Get-Date).ToUniversalTime().AddHours(0).ToString("yyyy-MM-ddTHH:mm:ssZ")
$timeFilter = (Get-Date).ToUniversalTime().AddHours($reportingDelay).ToString("yyyy-MM-ddTHH:mm:ssZ")

$callRecords = (Invoke-MgGraphRequest -Method Get "https://graph.microsoft.com/v1.0/communications/callRecords?`$filter=participants_v2/any(p:p/id eq '$($resourceAccountId)') and startDateTime lt $timeFilter" -ContentType "application/json" -OutputType PSObject).value

$callRecords = $callRecords | Where-Object { $_.type -eq "groupCall" } | Sort-Object -Property startDateTime

Write-Host "The report will include the following call records:" -ForegroundColor Cyan

$callRecords | Format-Table id, version, type, modalities, startDateTime, endDateTime, @{Name="CallerNumber";Expression={$_.organizer_v2.id.Substring(0,5)}}

$callSummaries = @()

$callRecordCounter = 1

# foreach ($callRecord in $callRecords | Where-Object { $_.startDateTime -gt (Get-Date -Date "17.05.2025").ToUniversalTime() }) {
foreach ($callRecord in $callRecords) {

    $callId = $callRecord.id

    Write-Host "Processing call record ID: $($callId) - Record $($callRecordCounter)/$($callRecords.Count)" -ForegroundColor Cyan

    $callRecordCounter ++

    $call = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/communications/callRecords/$callId" -ContentType "application/json" -OutputType PSObject

    $callOrganizerId = $call.organizer_v2.id

    $callOrganizer = $call.organizer_v2.id

    if ($callOrganizer -notmatch "^\+") {

        $callOrganizer = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/users/$($callOrganizer)" -ContentType "application/json").displayName

        $callerIsInternalCaller = $true

    }

    else {

        $callerIsInternalCaller = $false

    }

    $sessions = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/communications/callRecords/$($call.id)/sessions" -OutputType PSObject).value

    $calleeSideParticipants = $sessions.callee.identity.user
    $callerSideParticipants = $sessions.caller.identity.user

    $callStartDateTime = (Get-Date -Date $call.startDateTime).ToUniversalTime()
    $callEndDateTime = (Get-Date -Date $call.endDateTime).ToUniversalTime()

    $fromTime = $callStartDateTime.AddMinutes(-1).ToString("yyyy-MM-ddTHH:mm:ssZ")
    $toTime = $callEndDateTime.AddMinutes(1).ToString("yyyy-MM-ddTHH:mm:ssZ")

    $pstnCalls = (Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/v1.0/communications/callRecords/getPstnCalls(fromDateTime=$($fromTime),toDateTime=$($toTime))" -ContentType "application/json" -OutputType PSObject).value

    if (!$pstnCalls) {

        Write-Host "No matching PSTN call found. Disregarding call record."

    }

    $pstnCalls = $pstnCalls | Where-Object { $_.callType -eq "ucap_in" -or $_.callType -eq "oc_ucap_in" -or $_.callType -eq "ByotInUcap" }

    if ($pstnCalls.Count -gt 1) {

        # Check for matching call id for Direct Routing and Operator Connect calls
        $matchingPstnCall = $pstnCalls | Where-Object { $_.callId -eq $call.id }

        # If Calling Plan call
        if (-not $matchingPstnCall) {

            if ($pstnCalls.callerNumber -match "\*") {

                $matchingPstnCall = $pstnCalls | Where-Object { $($callOrganizer).Replace('+', '') -like "*$($_.callerNumber.Replace('*', '').Replace('+',''))*" -and $_.userId -in $sessions.caller.identity.user.id }

            }

            else {

                $matchingPstnCall = $pstnCalls | Where-Object { $_.callerNumber -eq $callOrganizer }

            }

        }

        if ($matchingPstnCall.Count -gt 1) {

            $matchingPstnCall = $matchingPstnCall | Sort-Object -Property callerNumber -Unique

        }

    }

    else {

        $matchingPstnCall = $pstnCalls

    }

    if (!$matchingPstnCall) {

        Write-Output "Internal call, no matching PSTN call"
        
        $topLevelResourceAccountSession = ($sessions | Sort-Object startDateTime | Select-Object -First 1).caller.identity.user

        if (!$topLevelResourceAccountSession.displayName) {

            $topLevelResourceAccount = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($topLevelResourceAccountSession.id)" -ContentType "application/json" -OutputType PSObject)

            $topLevelResourceAccountId = $topLevelResourceAccount.id
            $topLevelResourceAccount = $topLevelResourceAccount.displayName

        }

        else {

            $topLevelResourceAccount = $topLevelResourceAccountSession.displayName
            $topLevelResourceAccountId = $topLevelResourceAccountSession.id

        }


        $matchingPstnCall = @{ userId = $topLevelResourceAccountId }

        $resourceAccountNumber = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($topLevelResourceAccountId)" -ContentType "application/json").businessPhones

        $calleeNumber = $resourceAccountNumber[0]

    }

    else {

        $topLevelResourceAccount = $matchingPstnCall.userDisplayName
        $topLevelResourceAccountId = $matchingPstnCall.userId

        $calleeNumber = $matchingPstnCall.calleeNumber

    }

    $answeredSessions = @()

    $agentSessions = $sessions | Where-Object { $_.caller.identity.user.displayName -ne $null -and $_.caller.identity.user.id -ne $callOrganizerId }

    $agentSessions | Sort-Object -Property startDateTime | ft startDateTime, endDateTime, @{Name="duration";Expression={($_.endDateTime - $_.startDateTime).TotalSeconds}}, `
        @{Name="CallerSideUsers";Expression={$_.caller.identity.user.displayName}}

    foreach ($session in $agentSessions) {

        if ($null -ne $session.caller.identity.user.displayName -and $session.callee.userAgent.role -notin @("skypeForBusinessAutoAttendant", "skypeForBusinessCallQueues")) {

            # Azure Automation retrieves the session start and end date times as strings
            $sessionEndDateTime = (Get-Date -Date $session.endDateTime).ToUniversalTime()
            $sessionStartDateTime = (Get-Date -Date $session.startDateTime).ToUniversalTime()

            $sessionDuration = $($sessionEndDateTime) - $($sessionStartDateTime)

            if ($sessionStartDateTime -eq $sessionEndDateTime) {

                Write-Output "Session duration: $sessionDuration (no answered calls detected)"

            }

            else {

                Write-Output "Session duration: $sessionDuration"

                $answeredSessions += $session

            }

        }

    }

    $voiceAppSessions = $callerSideParticipants | Where-Object { $_.displayName -eq $null -and $_.id -notin @("9e133cac-5238-4d1e-aaa0-d8ff4ca23f4e", $($matchingPstnCall.userId)) }

    if (!$voiceAppSessions.id) {

        $voiceAppSessionsWithoutTopLevelResourceAccount = $calleeSideParticipants | Where-Object { $_.id -notin @("9e133cac-5238-4d1e-aaa0-d8ff4ca23f4e", $($matchingPstnCall.userId)) }

        if ($voiceAppSessionsWithoutTopLevelResourceAccount.Id.Count -gt 1) {

            $voiceAppSessionsStartTimes = ($sessions | Where-Object { $_.callee.identity.user.id -in $voiceAppSessionsWithoutTopLevelResourceAccount.id } | Sort-Object -Property startDateTime)[-1]

            $finalResourceAccount = $voiceAppSessionsStartTimes.callee.identity.user.displayName
            $finalResourceAccountId = $voiceAppSessionsStartTimes.callee.identity.user.id

            if ($finalResourceAccountId -and !$finalResourceAccount) {

                $finalResourceAccount = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($finalResourceAccountId)" -ContentType "application/json").displayName

            }

        }

        else {

            $finalResourceAccount = $voiceAppSessionsWithoutTopLevelResourceAccount.displayName
            $finalResourceAccountId = $voiceAppSessionsWithoutTopLevelResourceAccount.id

            if ($finalResourceAccountId -and !$finalResourceAccount) {

                $finalResourceAccount = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($finalResourceAccountId)" -ContentType "application/json").displayName

            }

        }

    }

    else {

        $finalResourceAccount = ($calleeSideParticipants | Where-Object { $_.id -eq $voiceAppSessions.id }).displayName
        $finalResourceAccountId = ($calleeSideParticipants | Where-Object { $_.id -eq $voiceAppSessions.id }).id

        if ($finalResourceAccountId -and !$finalResourceAccount) {

            $finalResourceAccount = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($finalResourceAccountId)" -ContentType "application/json").displayName

        }

    }

    Write-Output "Top level voice app (resource account) name: $topLevelResourceAccount"
    Write-Output "Final voice app (resource account) name: $finalResourceAccount"

    $callerNumber = $callOrganizer

    # Check agent sessions to determine if call was answered

    # Call never went into a queue
    if (!$agentSessions) {

        Write-Output "Call $($call.id) from $($callOrganizer) to resource account $($topLevelResourceAccount) was not forwarded to any call queue."

        Write-Output "Call $($call.id) from $($callOrganizer) to resource account $($topLevelResourceAccount) was not answered by an agent"

        $answeredSessions = $null

    }

    # Call went into a queue
    else {

        $answeredSessions = @()
        $missedSessions = @()

        foreach ($session in $agentSessions) {

            # Azure Automation retrieves the session start and end date times as strings
            $sessionEndDateTime = (Get-Date -Date $session.endDateTime).ToUniversalTime()
            $sessionStartDateTime = (Get-Date -Date $session.startDateTime).ToUniversalTime()

            $sessionDuration = $($sessionEndDateTime) - $($sessionStartDateTime)

            if ($sessionStartDateTime -eq $sessionEndDateTime) {

                $missedSessions += $session

                Write-Output "Session duration: $sessionDuration"

            }

            else {

                $answeredByAgent = $session.caller.identity.user.displayName | Where-Object { $_ -ne $null }

                Write-Output "Session duration: $sessionDuration"

                $answeredSessions += $session

            }

        }

    }

    if ($answeredSessions) {

        if ($answeredSessions.Count -gt 1) {

            $answeredSession = ($answeredSessions | Sort-Object -Property startDateTime -Descending)[-1]

        }

        else {

            $answeredSession = $answeredSessions

        }

        $answeredSessionStartDateTime = (Get-Date -Date $answeredSession.startDateTime).ToUniversalTime()
        $answeredSessionEndDateTime = (Get-Date -Date $answeredSession.endDateTime).ToUniversalTime()

        $timeInQueue = ($callStartDateTime - $answeredSessionStartDateTime).ToString("hh\:mm\:ss")

        $netCallDuration = ($answeredSessionEndDateTime - $answeredSessionStartDateTime).ToString("hh\:mm\:ss")

        Write-Output "Call $($call.id) from $($callOrganizer) to resource account $($topLevelResourceAccount) was answered by an agent: $answeredByAgent"

        $result = "Answered"

        $callDuration = ($callEndDateTime - $callStartDateTime).ToString("hh\:mm\:ss")

        $answeredPlatform = $answeredSession.caller.userAgent.platform

    }

    else {

        Write-Output "Call $($call.id) from $($callOrganizer) to resource account $($topLevelResourceAccount) was not answered by an agent"

        $result = "Missed"

        $answeredByAgent = $null

        $callDuration = ($callEndDateTime - $callStartDateTime).ToString("hh\:mm\:ss")

        $timeInQueue = $callDuration

        $netCallDuration = "00:00:00"

        $answeredPlatform = $null

    }

    $callSummary = [PSCustomObject]@{
        IsInternalCall = $callerIsInternalCaller
        CallType = $call.type
        CallerNumber = $callerNumber
        CalledVoiceApp = $topLevelResourceAccount
        AllInvolvedVoiceApps = (($sessions | Where-Object { $_.callee.userAgent.Role -in @("skypeForBusinessCallQueues", "skypeForBusinessAutoAttendant") -and $_.callee.identity.user.displayName -ne "" }).callee.identity.user.displayName | Sort-Object startDateTime | Sort-Object -Unique) -join "; "
        FinalVoiceApp = $finalResourceAccount
        CalledNumber = $calleeNumber
        StartDateTime = $callStartDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
        EndDateTime = $callEndDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
        Result = $result
        AnsweredBy = $answeredByAgent
        Platform = $answeredPlatform
        CallDuration = $callDuration
        NetCallDuration = $netCallDuration
        QueueDuration = $timeInQueue
        CallId = $call.id
        CallRecordVersion = $call.version
        CallRecordLastModifiedDateTime = (Get-Date -Date $call.lastModifiedDateTime).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    }

    Write-Host "Call summary:" -ForegroundColor Cyan

    # $callSummary

    $callSummaries += $callSummary

}

$callSummaries = $callSummaries | Sort-Object -Property StartDateTime

# $callSummaries | Export-Csv -Path "C:\Temp\CallSummariesAll-$userId.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ";"

$callSummaries | Out-GridView -Title "Call Summaries"