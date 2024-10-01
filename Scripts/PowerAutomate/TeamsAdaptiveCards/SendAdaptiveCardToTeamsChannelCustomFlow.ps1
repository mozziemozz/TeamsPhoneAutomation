$flowTriggerUri = ""

$flowHeaders = @{
    "TriggerSecret" = "OnlyICanTriggerThisFlow"
    "Channel" = "Azure"
}

$adaptiveCard = @"
{
    "type": "AdaptiveCard",
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "version": "1.3",
    "body": [
        {
            "type": "TextBlock",
            "text": "Hello Azure!",
            "wrap": true,
            "color": "Accent",
            "size": "ExtraLarge"
        }
    ]
}
"@

Invoke-RestMethod -Uri $flowTriggerUri -Method Post -Headers $flowHeaders -Body $adaptiveCard -ContentType "application/json"
