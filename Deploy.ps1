#requires -Version 5.1

$powerShellVersion = $PSVersionTable.PSVersion.Major

if ($powerShellVersion -ne 5) {

    powershell.exe -Version 5.1 -File $MyInvocation.MyCommand.Definition

}

if (!(Test-Path -Path C:\Temp)) {

    New-Item -Path C:\Temp -ItemType Directory

}

$repoPath = $MyInvocation.InvocationName | Split-Path -Parent

Set-Content -Path C:\Temp\RepoPath.txt -Value $repoPath

. .\Setup\Setup.ps1