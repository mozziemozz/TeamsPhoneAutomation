<#
    .SYNOPSIS
    This script validates dynamic membership rules for a group in Microsoft Entra.

    .DESCRIPTION
    The script reads a membership rule and a list of user IDs from specified paths.
    It then evaluates each user against the membership rule and generates a report.

    .PARAMETER MemberShipRulePath
    The path to the file containing the dynamic membership rule.

    .PARAMETER UserIdsPath
    The path to the file containing the user IDs to be evaluated.

    .EXAMPLE
    PS> .\Validate-DynamicEntraMemberShipRule.ps1 -MemberShipRulePath "path\to\rule.sql" -UserIdsPath "path\to\userids.txt"

    .NOTES
    Author: Martin Heusser
    Date: 2024-03-04
    Version: 1.0.1

    .LINK
    https://heusser.pro
    https://github.com/sponsors/mozziemozz
    https://github.com/mozziemozz/TeamsPhoneAutomation

#>

$localRepoPath = git rev-parse --show-toplevel
$localTestPath = "./.local"

# Rest of your script...


$localRepoPath = git rev-parse --show-toplevel
$localTestPath = "./.local"

Connect-MgGraph

Import-Module Microsoft.Graph.Beta.Groups

function Validate-DynamicEntraMemberShipRule {
    param (
        [string][Parameter(Mandatory = $true)]$MemberShipRulePath,
        [string[]][Parameter(Mandatory = $true)]$UserIdsPath
    )

    $membershipRule = Get-Content -Path $MemberShipRulePath | Out-String

    $userIds = Get-Content -Path $UserIdsPath

    $results = @()

    foreach ($userId in $userIds) {

        $params = @{

            MemberId       = $userId
            MembershipRule = $membershipRule
        }

        $evaluation = Test-MgBetaGroupDynamicMembershipRule -BodyParameter $params

        $userProperties = Get-MgUser -UserId $userId -Property City, UserPrincipalName, DisplayName, JobTitle

        $userDetails = New-Object -TypeName psobject

        $userDetails | Add-Member -MemberType NoteProperty -Name UserId -Value $userId
        $userDetails | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $userProperties.UserPrincipalName
        $userDetails | Add-Member -MemberType NoteProperty -Name DisplayName -Value $userProperties.DisplayName
        $userDetails | Add-Member -MemberType NoteProperty -Name JobTitle -Value $userProperties.JobTitle
        $userDetails | Add-Member -MemberType NoteProperty -Name Result -Value $evaluation.MembershipRuleEvaluationResult
        $userDetails | Add-Member -MemberType NoteProperty -Name City -Value $userProperties.City

        $results += $userDetails

    }

    $results | Format-Table -AutoSize

    $date = Get-Date -Format "yyyy-MM-dd-HH-mm"

    $results | Export-Csv -Path "$localTestPath/DynamicGroupEvaluationReport-$date.csv" -Delimiter ";" -Encoding utf8 -Force

}

Validate-DynamicEntraMemberShipRule -MemberShipRulePath "$localRepoPath/Scripts/EntraID/DynamicGroupResources/DynamicUserGroupRule.sql" -UserIdsPath "$localTestPath/DynamicUserGroupTestMembers.txt"

