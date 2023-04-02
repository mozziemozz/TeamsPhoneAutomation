# List source: https://countrycode.dev/docs#/Phone%20calls/get_calls_api_calls_get

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

    if ($country.Count -gt 1) {

        # In rare cases the prefix table can return multiple matching countries. E.g. US & CA. By default, the last matching country will be chosen. Add more entries with and digits to the prefix to avoid this problem.
        $country = $country[-1]

    }
    
    $voiceRoutingPolicy = ($prefixTable | Where-Object {$_.Prefix -eq $countryLookupResult}).VoiceRoutingPolicy

    return $country
    
}