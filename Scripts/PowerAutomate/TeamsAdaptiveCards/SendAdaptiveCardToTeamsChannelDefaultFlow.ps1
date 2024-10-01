$adaptiveCard = @"
{
  "type": "message",
  "attachments": [
    {
      "contentType": "application/vnd.microsoft.card.adaptive",
      "contentUrl": null,
      "content": {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.3",
        "body": [
          {
            "type": "TextBlock",
            "text": "Hello World 1!",
            "wrap": true,
            "color": "Accent",
            "size": "ExtraLarge"
          }
        ]
      }
    },
    {
      "contentType": "application/vnd.microsoft.card.adaptive",
      "contentUrl": null,
      "content": {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.3",
        "body": [
          {
            "type": "TextBlock",
            "text": "Hello World 2!",
            "wrap": true,
            "color": "Accent",
            "size": "ExtraLarge"
          }
        ]
      }
    }
  ]
}
"@

$flowTriggerUri = "https://prod-211.westeurope.logic.azure.com:443/workflows/34bc34e1ad8042fe97d0849ac998019c/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=4eQs-pqtvGG6mAY4lXms0OTnKembeDFebfwB4ekh5tM"

Invoke-RestMethod -Method Post -Uri $flowTriggerUri -Body $adaptiveCard -ContentType "application/json"