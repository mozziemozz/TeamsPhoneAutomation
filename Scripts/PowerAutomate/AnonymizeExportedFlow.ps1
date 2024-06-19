<#

.SYNOPSIS
    Anonymize exported flow files

.DESCRIPTION
    Anonymize exported flow files

.NOTES
    Author:     Martin Heusser | MVP
    Version:    1.0.0
    Changelog:  2023-12-06: Initial release

.Example
    .\AnonymizeExportedFlow.ps1

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $FilePath,
    [Parameter(Mandatory = $true)]
    [string]
    $OldDomain,
    [Parameter(Mandatory = $true)]
    [string]
    $NewDomain,
    [Parameter(Mandatory = $false)]
    [string]
    $OldSharePointUrl,
    [Parameter(Mandatory = $false)]
    [string]
    $NewSharePointUrl
)

$flowFiles = Get-ChildItem -Path $filePath -Recurse -File -Include "definition.json"

$guidMatch = '"([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})"'
$anonymousGuid = '"00000000-0000-0000-0000-000000000000"'

$channelIdMatch = '(19:[0-9a-fA]{32}@thread\.tacv2)'
$anonymousChannelId = '19:00000000000000000000000000000000@thread.tacv2'

foreach ($file in $flowFiles) {

    $fileContent = Get-Content -Path $file.FullName

    # Replace all occurrences of the old domain in the entire content
    $modifiedContent = $fileContent -replace $oldDomain, $newDomain

    # Replace all occurrences of the GUID pattern in the entire content
    $modifiedContent = $modifiedContent -replace $guidMatch, $anonymousGuid

    # Replace all occurrences of the channel ID pattern in the entire content
    $modifiedContent = $modifiedContent -replace $channelIdMatch, $anonymousChannelId

    # Replace all occurrences of the old SharePoint URL in the entire content
    $modifiedContent = $modifiedContent -replace $oldSharePointUrl, $newSharePointUrl

    # If you want to save the changes back to the file
    $modifiedContent | Set-Content -Path $file.FullName

}