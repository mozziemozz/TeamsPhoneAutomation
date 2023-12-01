<#
    .SYNOPSIS
    Clears the Teams client cache on Windows machines.

    .DESCRIPTION
    Clears the Teams client cache for whichever Teams version is currently in use while retaining the custom backgrounds.

    Authors:            Martin Heusser | MVP
    Version:            1.0.1
    Changelog:          2023-12-01: Initial release

    .PARAMETER Name
        None.

    .INPUTS
        None.

    .OUTPUTS
        None.

    .EXAMPLE
    . .\ClearTeamsClientCache.ps1

#>

if (!(Test-Path -Path "$env:APPDATA\Microsoft\Teams") -and !(Test-Path -Path "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe")) {

    Write-Host "Cache folders for either Teams version could not be found." -ForegroundColor Magenta

    Read-Host "Press any key to exit..."
    Exit

}

else {

    $teamsProcesses = Get-Process -Name *Teams*

    if ($teamsProcesses) {

        Write-Host "Stopping Teams processes..." -ForegroundColor Magenta

        foreach ($process in $teamsProcesses) {

            Stop-Process -Id $process.Id -ErrorAction SilentlyContinue

        }

        Start-Sleep -Seconds 10

        switch ($teamsProcesses[0].ProcessName) {
    
            "ms-teams" {

                Write-Host "New Teams is in use" -ForegroundColor Cyan

                $backgrounds = Get-ChildItem -Path "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds\Uploads"

                $tempFiles = @()

                foreach ($background in $backgrounds) {

                    $tempFiles += @{
                        FileName    = $background.Name
                        FullName    = $background.FullName
                        FileContent = [System.IO.File]::ReadAllBytes($background.FullName)
                    }

                }

                Remove-Item -Path "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams" -Recurse -Force

                New-Item -Path "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams" -ItemType Directory

                New-Item -Path "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds" -ItemType Directory

                New-Item -Path "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds\Uploads" -ItemType Directory

                foreach ($tempFile in $tempFiles) {

                    [System.IO.File]::WriteAllBytes($tempFile.FullName, $tempFile.FileContent)

                }

                Start-Process ms-teams.exe

            }

            "Teams" {

                Write-Host "Old Teams is in use" -ForegroundColor Cyan

                $backgrounds = Get-ChildItem -Path "$env:APPDATA\Microsoft\Teams\Backgrounds\Uploads"

                $tempFiles = @()

                foreach ($background in $backgrounds) {

                    $tempFiles += @{
                        FileName    = $background.Name
                        FullName    = $background.FullName
                        FileContent = [System.IO.File]::ReadAllBytes($background.FullName)
                    }

                }

                Remove-Item -Path "$env:APPDATA\Microsoft\Teams" -Recurse -Force

                New-Item -Path "$env:APPDATA\Microsoft\Teams" -ItemType Directory

                New-Item -Path "$env:APPDATA\Microsoft\Teams\Backgrounds" -ItemType Directory

                New-Item -Path "$env:APPDATA\Microsoft\Teams\Backgrounds\Uploads" -ItemType Directory

                foreach ($tempFile in $tempFiles) {

                    [System.IO.File]::WriteAllBytes($tempFile.FullName, $tempFile.FileContent)

                }

                Set-Location "$env:LOCALAPPDATA\Microsoft\Teams"
                Start-Process -File "$env:LOCALAPPDATA\Microsoft\Teams\Update.exe" -ArgumentList '--processStart "Teams.exe"'

            }

            Default {
            }

        }

    }

    else {

        Write-Host "Neither version of Teams is running. Please start Teams and run this script again." -ForegroundColor Magenta

        $checkIfNewTeamsIsInstalled = (Test-Path -Path "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe")

        if ($checkIfNewTeamsIsInstalled -eq $true) {

            Write-Host "New Teams is installed" -ForegroundColor Cyan
            
            $oldTeamsLastAccessTime = Get-ItemProperty -Path "$env:LOCALAPPDATA\Microsoft\Teams\Current\Teams.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty LastAccessTimeUtc

            $newTeamsLogs = (Get-ChildItem -Path "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs" -File -Filter "MSTeams_*.log")[-1]

            $newTeamsLogContent = Get-Content -Path $newTeamsLogs.FullName | Out-String

            $newTeamsVersionPattern = "LatestVersion:\s*(\d+(\.\d+)*)"
            $match = [regex]::Match($newTeamsLogContent, $newTeamsVersionPattern)

            $versionNumber = $match.Groups[1].Value

            $newTeamsLastAccessTime = Get-ItemProperty -Path "C:\Program Files\WindowsApps\MSTeams_$($versionNumber)_x64__8wekyb3d8bbwe\ms-teams.exe" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty LastAccessTimeUtc

            if (!$oldTeamsLastAccessTime -and !$newTeamsLastAccessTime) {

                Write-Host "Teams is not installed. Could not get last used time of either Teams version." -ForegroundColor Magenta

                $lastUsedVersion = "Teams is not installed."

            }

            elseif ($newTeamsLastAccessTime -gt $oldTeamsLastAccessTime) {

                $lastUsedVersion = "New Teams"

            }

            else {
            
                $lastUsedVersion = "Old Teams"

            }

        }

        else {

            Write-Host "New Teams is not installed"

            $lastUsedVersion = "Old Teams"

        }

        Write-Host "Last used Teams version: $lastUsedVersion" -ForegroundColor Cyan

    }

}