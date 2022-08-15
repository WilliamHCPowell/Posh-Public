#
# script to schedule a daily check of submission activity
#

param ($JobName = "SubmissionsCheck")
Import-Module BPJobs -Verbose

Connect-BPScheduledJobManager

#
# passing a parameter here - see https://stackoverflow.com/questions/16347214/pass-arguments-to-a-scriptblock-in-powershell 
$scriptBlock = {
    cd C:\Projects\Author
    .\Run-DailySubmissionsCheck.ps1 -JobName $args[0]
}

New-BPScheduledJob -JobName $JobName `
                   -ScriptBlock $scriptBlock `
                   -TriggerType Days1 `
                   -StartTime "01:15" `
                   -ArgumentList @($JobName)

Write-Host "Done"
