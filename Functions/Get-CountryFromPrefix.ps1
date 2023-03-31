function Get-CountryFromPrefix {
    param (
        
    )

    switch ($localTestMode) {
        $true {
    
            # Local Environment
            
            $prefixTable = Get-Content -Path .\Resources\CountryLookupTable.json | ConvertFrom-Json
            
        }
    
        $false {
    
            # Azure Automation
            
            $prefixTable = (Get-AutomationVariable -Name "TeamsPhoneNumberOverview_CountryLookupTable").Replace("'","") | ConvertFrom-Json
    
        }
        Default {}
    }

    $countryLookupResults = @()

    foreach ($prefix in $prefixTable) {

        if ($phoneNumber.StartsWith($prefix.Prefix)) {

            $countryLookupResults += $prefix.Prefix

        }

    }
    
    if ($countryLookupResults.Count -gt 1) {
    
        $countryLookupResult = $countryLookupResults | Sort-Object length -Descending | Select-Object -first 1
    
    }
    
    else {
    
        $countryLookupResult = $countryLookupResults
    
    }
    
    $country = ($prefixTable | Where-Object {$_.Prefix -eq $countryLookupResult}).Country
    $voiceRoutingPolicy = ($prefixTable | Where-Object {$_.Prefix -eq $countryLookupResult}).VoiceRoutingPolicy
    # $country = (Get-Culture).TextInfo.ToTitleCase($country.ToLower())

    return $country
    
}