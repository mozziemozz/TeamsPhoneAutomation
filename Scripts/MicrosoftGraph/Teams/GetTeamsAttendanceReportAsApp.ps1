<#

    .SYNOPSIS
    Get Teams attendance reports as an app.

    .DESCRIPTION
    This script gets attendance report of Teams meetings as an Entra ID app without a signed in user.

    Author:             Martin Heusser
    Version:            1.0.0
    Changelog:          2024-10-05: Initial release
    
    Website:            https://heusser.pro
    Blog Post:          https://heusser.pro/p/get-microsoft-teams-meeting-attendance-report-through-graph-api-lhpctbnzht7z/
    Sponsor Project:    https://github.com/sponsors/mozziemozz
    Buy me Coffee:      https://buymeacoffee.com/martin.heusser

    This script inlcudes an example for all supported filter queries to get users by business or mobile phone number

    .EXAMPLE
    .\GetTeamsAttendanceReportAsApp.ps1

    .NOTES
    Make sure to add your own Ids bellow the function.


#>

#requires -Module MicrosoftTeams, Microsoft.Graph.CloudCommunications

function Get-TeamsAttendanceReportAsApp {
    param(
        [Parameter(Mandatory = $true)][String]$TenantId,
        [Parameter(Mandatory = $true)][String]$AppId,
        [Parameter(Mandatory = $true)][String]$AppSecret,
        [Parameter(Mandatory = $true)][String]$UserId,
        [Parameter(Mandatory = $true)][String]$JoinMeetingId,
        [Parameter(Mandatory = $false)][String]$ApplicationAccessPolicyName = "AccessAttendanceReportAsApp"
    )

    $checkTeamsSession = Get-CsTenant -ErrorAction SilentlyContinue

    if (!$checkTeamsSession) {

        Write-Host "Sign into Teams PowerShell in Browser..." -ForegroundColor DarkMagenta

        Connect-MicrosoftTeams

    }

    $checkApplicationAccessPolicy = Get-CsApplicationAccessPolicy -Identity $applicationAccessPolicyName -ErrorAction SilentlyContinue

    if (!$checkApplicationAccessPolicy) {

        Write-Host "Application Access Policy '$applicationAccessPolicyName' not found. Creating a new one..." -ForegroundColor Yellow

        New-CsApplicationAccessPolicy -Identity $applicationAccessPolicyName -AppIds $AppId -Description "Access Meeting Data as Entra ID app on behalf of user"

    }

    $checkCsOnlineUser = Get-CsOnlineUser -Identity $UserId -ErrorAction SilentlyContinue

    if ($checkCsOnlineUser.ApplicationAccessPolicy.Name -notcontains $applicationAccessPolicyName) {

        Write-Host "Application Access Policy '$applicationAccessPolicyName' not assigned to user '$UserId'. Assigning..." -ForegroundColor Yellow

        Grant-CsApplicationAccessPolicy -PolicyName $applicationAccessPolicyName -Identity $UserId

        do {

            Start-Sleep -Seconds 10

            $checkCsOnlineUser = Get-CsOnlineUser -Identity $UserId -ErrorAction SilentlyContinue

        } until (
            $checkCsOnlineUser.ApplicationAccessPolicy.Name -contains $applicationAccessPolicyName
        )

        Start-Sleep -Seconds 30

        Write-Host "Application Access Policy '$applicationAccessPolicyName' assigned to user '$UserId'." -ForegroundColor Green

    }

    $onlineMeeting = Get-MgUserOnlineMeeting -UserId $userId -Filter "joinMeetingIdSettings/joinMeetingId eq '$joinMeetingId'"

    if (!$onlineMeeting) {

        Write-Host "Online meeting with join meeting id '$joinMeetingId' not found. Try again in a couple of seconds..." -ForegroundColor Red

        Read-Host "Press any key to exit"

        exit

    }

    $attendanceReports = Get-MgUserOnlineMeetingAttendanceReport -UserId $userId -OnlineMeetingId $onlineMeeting.Id

    $attendanceReportOutput = @()

    foreach ($attendanceReport in $attendanceReports) {

        $attendanceReportSummary = @()

        $attendanceReportDetails = Get-MgUserOnlineMeetingAttendanceReport -UserId $userId -OnlineMeetingId $onlineMeeting.Id -MeetingAttendanceReportId $attendanceReport.Id -Expand "attendanceRecords"

        foreach ($attendanceReportRecord in $attendanceReportDetails.attendanceRecords) {

            $attendanceReportOutput += [pscustomobject]@{

                "Email"             = $attendanceReportRecord.emailAddress
                "DisplayName"       = $attendanceReportRecord.identity.displayName
                "Role"              = $attendanceReportRecord.role
                "DurationInSeconds" = $attendanceReportRecord.totalAttendanceInSeconds
                "Duration"          = [timespan]::FromSeconds($attendanceReportRecord.totalAttendanceInSeconds).ToString()
                "JoinDateTime"      = $attendanceReportRecord.attendanceIntervals.joinDateTime
                "LeaveDateTime"     = $attendanceReportRecord.attendanceIntervals.leaveDateTime
                "ReportId"          = $attendanceReport.id

            }

        }

        $meetingDuration = $attendanceReport.MeetingEndDateTime - $attendanceReport.MeetingStartDateTime

        $averageAttendanceTime = ($attendanceReportDetails.attendanceRecords | Measure-Object -Property totalAttendanceInSeconds -Average).Average

        $attendanceReportSummary += [pscustomobject]@{

            "MeetingSubject" = $onlineMeeting.Subject
            "Participants"   = $attendanceReport.TotalParticipantCount
            "AverageAttendanceTime" = [timespan]::FromSeconds($averageAttendanceTime).ToString()
            "StartTime"      = $attendanceReport.MeetingStartDateTime
            "EndTime"        = $attendanceReport.MeetingEndDateTime
            "MeetingDuration" = [timespan]::FromSeconds($meetingDuration.TotalSeconds).ToString()
            "ReportId"       = $attendanceReport.id

        }

        $attendanceReportSummary | Export-Csv -Path ".\$($onlineMeeting.Subject)-AttendanceReportSummary-$($attendanceReport.Id).csv" -NoTypeInformation -Encoding UTF8 -Delimiter ";"

    }

    $attendanceReportOutput | Export-Csv -Path ".\$($onlineMeeting.Subject)-AttendanceReport.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ";"

    Write-Host "Attendance report saved to .\$($onlineMeeting.Subject)-AttendanceReport.csv" -ForegroundColor Cyan

}

# Add your Ids here
$tenantId = "" # This is the tenant id
$appId = "" # This is the application id of the app you created with the OnlineMeetings.Read.All and OnlineMeetingArtifact.Read.All permissions

$applicationAccessPolicyName = "AccessTeamsAttendanceReportsAsApp" # This is the name of the application access policy you want to create

# Add your Ids here
$joinMeetingId = "" # This is the join meeting id of the meeting you want to get the attendance report for
$userId = "" # This is the user id of the meeting organizer you want to get the attendance report for

# You need to clone the whole repository to get the SecureCredsMgmt.ps1 file
# Import (dot source) SecureCredsMgmt functions
. .\Modules\SecureCredsMgmt.ps1

. Get-MZZSecureCreds -FileName "AccessTeamsMeetingDataAsAppDemo"
$appSecret = $passwordDecrypted

# Create new powershell credential object from app id (user name) and app secret (password)
$clientSecretCredential = New-Object System.Management.Automation.PSCredential ($appId, (ConvertTo-SecureString $appSecret -AsPlainText -Force))

# Connect to Graph
Connect-MgGraph -ClientSecretCredential $clientSecretCredential -TenantId $tenantId -NoWelcome

# Call the function to generate the attendance report
. Get-TeamsAttendanceReportAsApp -TenantId $tenantId -AppId $appId -AppSecret $appSecret -UserId $userId -JoinMeetingId $joinMeetingId -ApplicationAccessPolicyName $applicationAccessPolicyName