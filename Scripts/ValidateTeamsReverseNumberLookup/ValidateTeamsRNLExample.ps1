# Change these variables to match your environment
$adminUser = "user@domain.com"
$mfa = $false
$hideObjectId = $false

$lineURIs = @(

    "+41xxxxxxxxx",
    "+41xxxxxxxxx",
    "+41xxxxxxxxx"

)

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