#
# Scheduled Job management
#

#region Scheduled Job Management

function Connect-BPScheduledJobManager {
}

function New-BPTrigger ([ValidateSet('Days1', 'Hours1', 'Minutes15', 'Minutes1')]   [string]$TriggerType,
                          $StartTime) {
    # start from the current time
    $now = Get-Date

    $nowPlusSafetyMargin = $now.AddMinutes(2)

    $trigger = $null

    switch ($TriggerType) {
        "Days1" {
            $triggerStartTime = $StartTime
            $trigger = New-JobTrigger -Daily -At $triggerStartTime
        }
        "Hours1" {
            # set up a trigger, which is 20 seconds after the hour, not less than the safety margin into the future
            #
            # step back from now to the start of the day, plus 20 seconds
            $triggerStartTime = [datetime]::Parse( $now.ToString("yyyy/MM/dd") + " 00:00:20" )
            #
            # step forward in 15 minute steps until we have a time not less than the safety margin into the future
            while ($triggerStartTime -lt $nowPlusSafetyMargin) { $triggerStartTime = $triggerStartTime.AddHours(1) }

            $trigger = New-JobTrigger -Once `
                -At $triggerStartTime `
                -RepetitionInterval "01:00:00" `
                -RepeatIndefinitely
        }
        "Minutes15" {
            # set up a trigger, which is 20 seconds after the quarter hour, not less than the safety margin into the future
            #
            # step back from now to the start of the day, plus 20 seconds
            $triggerStartTime = [datetime]::Parse( $now.ToString("yyyy/MM/dd") + " 00:00:20" )
            #
            # step forward in 15 minute steps until we have a time not less than the safety margin into the future
            while ($triggerStartTime -lt $nowPlusSafetyMargin) { $triggerStartTime = $triggerStartTime.AddMinutes(15) }

            $trigger = New-JobTrigger -Once `
                -At $triggerStartTime `
                -RepetitionInterval (New-TimeSpan -Minutes 15) `
                -RepeatIndefinitely
        }
        "Minutes1" {
            # set up a trigger, which is 20 seconds after the quarter hour, not less than the safety margin into the future
            #
            # step back from now to the start of the day, plus 20 seconds
            $triggerStartTime = [datetime]::Parse( $now.ToString("yyyy/MM/dd") + " 00:00:20" )
            #
            # step forward in 15 minute steps until we have a time not less than the safety margin into the future
            while ($triggerStartTime -lt $nowPlusSafetyMargin) { $triggerStartTime = $triggerStartTime.AddMinutes(15) }

            $trigger = New-JobTrigger -Once `
                -At $triggerStartTime `
                -RepetitionInterval (New-TimeSpan -Minutes 1) `
                -RepeatIndefinitely
        }
    }
    $trigger
}

function New-BPScheduledJob ($JobName,
                               $ScriptBlock,
                               [ValidateSet('Days1', 'Hours1', 'Minutes15')]   [string]$TriggerType,
                               $StartTime,
                               $Credential,
                               $ArgumentList) {
    $trigger = New-BPTrigger -TriggerType $TriggerType -StartTime $StartTime

    $job = Get-ScheduledJob -Name $JobName -ErrorAction SilentlyContinue

    if ($job -eq $null) {
        if ($ArgumentList -eq $null) {
            $ArgumentList = @()
        }
        $RegisterParms = @{
            Name = $JobName;
            Trigger = $trigger;
            ScriptBlock = $ScriptBlock;
        }
        if ($ArgumentList -ne $null) {
            $RegisterParms["ArgumentList"] = $ArgumentList
        }
        if ($Credential -ne $null) {
            $RegisterParms["Credential"] = $Credential
        }
        $job = Register-ScheduledJob @RegisterParms
    }
    else {
        if ($job.JobTriggers.Count -eq 0) {
            Write-Host "Added trigger for Scheduled job $JobName"
            Add-JobTrigger -Trigger $trigger -Name $JobName
        }
    }
}

function Remove-BPScheduledJob ($JobName) {
    #
    # could disable them, but while we're getting the hang of scheduled jobs
    # just zap the lot
    Write-Host "Removing all triggers for scheduled job $JobName"
    Remove-JobTrigger -Name $JobName
    Unregister-ScheduledJob -Name $JobName
}

function Get-BPScheduledJob ($JobName,$ServerList) {
}

Export-ModuleMember Connect-BPScheduledJobManager,New-BPScheduledJob,Remove-BPScheduledJob,Get-BPScheduledJob

#endregion
