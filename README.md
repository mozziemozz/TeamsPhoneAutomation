# Teams Phone Automation

## Teams Phone Number Management

Author: Martin Heusser | M365 Apps & Services MVP

Please check out the accompanying [blog post](https://medium.com/@mozzeph/teams-phone-number-management-on-a-budget-e25d53f65caf).

The readme will be updated at a later time.

## Teams Managament Module

The module located in `.\Modules\TeamsPS.psm1` contains some useful functions to help with day to day Teams administrative tasks.

### Import the Module

```powershell
$localRepoPath = git rev-parse --show-toplevel
Import-Module "$localRepoPath/Modules/TeamsPS.psm1" -Force
```