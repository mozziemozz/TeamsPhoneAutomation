function Get-CountryFromPrefix {
    param (
        
    )

    $prefixTable = Get-Content -Path .\Resources\CountryLookupTable.json | ConvertFrom-Json

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
    $country = (Get-Culture).TextInfo.ToTitleCase($country.ToLower())

    return $country
    
}