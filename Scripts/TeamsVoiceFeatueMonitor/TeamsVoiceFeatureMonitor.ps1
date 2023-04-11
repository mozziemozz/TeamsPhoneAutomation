# Set to true if script is executed locally
$localTestMode = $true

switch ($localTestMode) {
    $true {

        # Local Environment

        # Import external functions
        . .\Functions\Connect-MsTeamsServicePrincipal.ps1

        # Import variables
        $TenantId = Get-Content -Path .\.local\TenantId.txt
        $AppId = Get-Content -Path .\.local\AppId.txt
        $AppSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR((Get-Content -Path .\.local\AppSecret.txt | ConvertTo-SecureString))) | Out-String

        $ExistingFlightedFeatures = (Get-Content -Path .\.local\FlightedFeatures.txt).Replace("'","") | ConvertFrom-Json

        $WebhookUrl = Get-Content -Path .\.local\WebhookURL.txt
        
    }

    $false {

        # Azure Automation

        # Import external functions
        . .\Connect-MsTeamsServicePrincipal.ps1
        . .\Get-CsOnlineNumbers.ps1

        # Import variables        
        $TenantId = Get-AutomationVariable -Name "TeamsPhoneNumberOverview_TenantId"
        $AppId = Get-AutomationVariable -Name "TeamsPhoneNumberOverview_AppId"
        $AppSecret = Get-AutomationVariable -Name "TeamsPhoneNumberOverview_AppSecret"

        $ExistingFlightedFeatures = (Get-AutomationVariable -Name "FlightedFeatures").Replace("'","") | ConvertFrom-Json

        $WebhookUrl = Get-AutomationVariable -Name "WebhookUrl"

    }
    Default {}
}

. Connect-MsTeamsServicePrincipal -TenantId $TenantId -AppId $AppId -AppSecret $AppSecret

$flightedFeatures = (Get-CsAutoAttendantTenantInformation).FlightedFeatures | ConvertTo-Json | ConvertFrom-Json

$newFeatures = $flightedFeatures | Where-Object {$ExistingFlightedFeatures -notcontains $_}

if ($newFeatures) {

    $newFeatureList = ""

    foreach ($newFeature in $newFeatures) {

        $newFeatureList += "\n\n$newFeature"

    }

    $discordMessage =@"
{
    "content": null,
    "embeds": [
      {
        "title": "New Teams Voice Feature Discovered!",
        "description": "New Features:$newFeatureList",
        "color": 42223
      }
    ],
    "attachments": []
  }
"@

    Invoke-RestMethod -uri $WebhookUrl -Method Post -body $discordMessage -ContentType 'application/json; charset=UTF-8'

    switch ($localTestMode) {
        $true {

            Set-Content -Path .\.local\FlightedFeatures.txt -Value "'$($flightedFeatures | ConvertTo-Json)'"

        }

        $false {

            Set-AutomationVariable -Name "FlightedFeatures" -Value "'$($flightedFeatures | ConvertTo-Json)'"

        }
        Default {}
    }

}