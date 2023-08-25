<#
    .SYNOPSIS
    Example script to validate Teams phone number assignments by using reverse number lookup function from this repo.

    .DESCRIPTION
    Author:             Martin Heusser
    Version:            1.0.0
    Sponsor Project:    https://github.com/sponsors/mozziemozz
    Website:            https://heusser.pro

    .EXAMPLE
    . .\ValidateTeamsRNLExample.ps1

#>

# Change these variables to match your environment
$adminUser = "user@domain.com"
$mfa = $false
$hideObjectId = $false

$lineURIs = @(

    "+41xxxxxxxxx",
    "+41xxxxxxxxx",
    "+41xxxxxxxxx"

)

# End of adjustable variables

# Import functions by dot sourcing
. .\Scripts\ValidateTeamsReverseNumberLookup\ValidateTeamsReverseNumberLookup.ps1

$results = @()

foreach ($lineURI in $lineURIs) {

    . Test-MZZTeamsLineURIAssignment -LineURI $lineURI -AdminUser $adminUser -MFA $mfa

    if ($assignmentIsValid) {

        if ($hideObjectId) {

            $objectId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    
        }

        else {

            $objectId = $reverseNumberLookup.objectId

        }

    }

    else {

        $objectId = "NotFound"

    }

    $resultDetails = [PSCustomObject]@{

        LineURI = $lineURI
        AssignmentIsValid = $assignmentIsValid
        ObjectId = $objectId

    }

    $results += $resultDetails

}

$results | Format-Table -AutoSize