Connect-MicrosoftTeams

$voices = Get-CsAutoAttendantSupportedLanguage 

$voices | fl id, DisplayName, @{Name = "Voices"; Expression = {($_.Voices | % { "$($_.Name) ($($_.Id))" }) -join ", "}}