function Get-CsOnlineNumbers {
    param (
        
    )

    $numberTypes = @(

        "CallingPlan",
        "OperatorConnect"

    )

    $allCsOnlineNumbers = @()

    foreach ($numberType in $numberTypes) {

        $csOnlineNumbers = Get-CsPhoneNumberAssignment -NumberType $numberType -Top 500

        Write-Host "Getting $numberType numbers..."

        if ($csOnlineNumbers) {

            Write-Host "The first result is $($csOnlineNumbers.Count) numbers..."
    
            if ($csOnlineNumbers.Count -ge 500) {
        
                $skipCounter = 500
        
                do {
        
                    Write-Host "Skipping the first $skipCounter numbers..."
        
                    $querriedNumbers = Get-CsPhoneNumberAssignment -IsoCountryCode $geoLocation.IsoCountryCode -Skip $skipCounter
        
                    Write-Host "Found $($querriedNumbers.Count) Numbers..."
        
                    $csOnlineNumbers += $querriedNumbers
        
                    $skipCounter += 500
        
                } until (
                    $querriedNumbers.Count -eq 0
                )
        
            }

        }
    
        Write-Host "Finished getting $numberType numbers."
    
        $allCsOnlineNumbers += $csOnlineNumbers
    
    }

    $allCsOnlineNumbers = $allCsOnlineNumbers | Sort-Object TelephoneNumber -Unique

    return $allCsOnlineNumbers
    
}