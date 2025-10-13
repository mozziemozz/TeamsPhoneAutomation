$auth = Get-Content -Path ".\Scripts\Teams\Chats\auth.json" | ConvertFrom-Json

$userId = "339a9570-62f6-4769-82f7-5d29a0954037" # User object id of meeting chat participant

$clientSecret = $auth.secret | ConvertTo-SecureString

$clientSecretCredentials = New-Object System.Management.Automation.PSCredential($auth.appId, $clientSecret)

Connect-MgGraph -TenantId $auth.tenantId -Credential $clientSecretCredentials -NoWelcome

Get-MgContext | Format-List AppName, ClientId, TenantId, AuthType, TokenCredentialType, Scopes

$mgUser = Get-MgUser -UserId $userId

$meetingChats = Get-MgUserChat -UserId $mgUser.Id -All -Filter "chatType eq 'meeting' and lastUpdatedDateTime gt $((Get-Date).AddDays(-30).ToString('yyyy-MM-ddTHH:mm:ssZ'))"

if ($meetingChats.Count -gt 1) {

    $meetingChatSelection = $meetingChats | Select-Object Topic, Id | Out-GridView -Title "Select Meeting Chat" -OutputMode Single

    $meetingChat = $meetingChats | Where-Object { $_.Id -eq $meetingChatSelection.Id }

}

else {

    $meetingChat = $meetingChats

}

$meetingChatMessagesEventDetails = @()

$meetingChatMessagesAPI = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/chats/$($meetingChat.Id)/messages" -OutputType PSObject

$meetingChatMessagesEventDetails += $meetingChatMessagesAPI.value

do {

    $meetingChatMessagesAPI = Invoke-MgGraphRequest -Method GET -Uri $meetingChatMessagesAPI.'@odata.nextLink' -OutputType PSObject

    $meetingChatMessagesEventDetails += $meetingChatMessagesAPI.value

} until (
    -not $meetingChatMessagesAPI.'@odata.nextLink'
)

$meetingChatMessagesEventDetails | Where-Object { $_.eventDetail.'@odata.type' -match 'callRecordingEventMessageDetail' } | Select-Object createdDateTime, `
@{Name = "CallRecordingDisplayName"; Expression = { $_.eventDetail.callRecordingDisplayName } }, `
@{Name = "EventDetail"; Expression = { $_.eventDetail.'@odata.type' } }, @{Name = "CallRecordingStatus"; Expression = { $_.eventDetail.callRecordingStatus } }, `
@{Name = "CallRecordingInitiator"; Expression = { (Get-MgUser -UserId $_.eventDetail.initiator.user.id | Select-Object -ExpandProperty UserPrincipalName) } }, `
@{Name = "CallId"; Expression = { $_.eventDetail.callId } }, `
@{Name = "CallRecordingDuration"; Expression = { $_.eventDetail.callRecordingDuration } }, `
# @{Name = "CallRecordingUrl"; Expression = { $_.eventDetail.callRecordingUrl } } | Out-GridView
@{Name = "CallRecordingUrl"; Expression = { if ($_.eventDetail.callRecordingUrl -notmatch "graph") { $_.eventDetail.callRecordingUrl } } } | Out-GridView -Title "Meeting Recording Log"