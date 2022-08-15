#
# script to plot submissions over time
#

param ($JobName = "Local",
       $SubmissionData = 'P:\Users\Bill\Documents\My Stories\Submissions\All Submissions.sonar3')

Add-Type -AssemblyName system.web

Import-Module ImportExcel

#region Setup
#
# The script can take hours to run on a large dataset
# We need to report progress.  For short-ish tasks, up to about 30s
# we simply need to use Write-Host to output timely status messages
#
# (we use Write-Progress to show progress of longer tasks)
#
$ScriptElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()

$lastt = 0

function reportPhaseComplete ([string]$description) {
    $t = $ScriptElapsedTime.Elapsed.TotalSeconds
    $phaset = [Math]::Floor(($t - $script:lastt) * 10) / 10
    write-host "Phase complete, taking $phaset seconds: $description"
    $script:lastt = $t
}

function reportScriptComplete ([string]$description) {
  $t = $ScriptElapsedTime.Elapsed.TotalSeconds
  $phaset = [Math]::Floor(($t) * 10) / 10
  write-host "Script complete, taking $phaset seconds: $description"
  $script:lastt = $t
}
#
# standard functions to find the directory in which the script is executing
# we'll use this info to read and write both cache files and reports
#
function Get-ScriptDirectory {
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
    Split-Path $Invocation.MyCommand.Path
}

$sdir = Get-ScriptDirectory
$infoColours = @{foreground="cyan"}
$warningColours = @{foreground="yellow"}
$errorColours = @{foreground="red"}
$debugColours = @{foreground="green"}
#endregion

#region File Management
#
# Having identified the current working directory, we can now set up paths for the
# various cache files and report files used by the script.
#
$SubmissionFile               = "$sdir\SubmissionReport.xlsx"              # These files are shared by many jobs
$MarketFile                   = "$sdir\Markets.csv"                        #
$MarketUpdateFile             = "$sdir\MarketsUpdate.csv"                  #
$MarketStatusFile             = "$sdir\MarketStatus.json"                  #
$AgentInfoFile                = "$sdir\Agents.csv"                         #
$RalanMarketsFile             = "$sdir\RalanMarkets.csv"
$associationMemberListFile    = "$sdir\AAAList.csv"
$individualAgentsListFile     = "$sdir\Agents.csv"
$individualAgentsSheetFile    = "$sdir\Agents.xlsx"
$NewSubmissionFile            = "$sdir\SubmissionReport2.xlsx"
#endregion

$SubmissionsText = Get-Content -LiteralPath $SubmissionData
$SubmissionsDoc = [xml]$SubmissionsText

$markets = $SubmissionsDoc.SONAR3.MARKETS.MARKET
$works = $SubmissionsDoc.SONAR3.WORKS.WORK
$submissions = $SubmissionsDoc.SONAR3.SUBMISSIONS.SUBMISSION

$worksHash = @{}

$total = New-Object PSObject |
    Add-Member NoteProperty Date         $null -PassThru |
    Add-Member NoteProperty Works        0     -PassThru |
    Add-Member NoteProperty Acceptances  0     -PassThru |
    Add-Member NoteProperty Rejections   0     -PassThru |
    Add-Member NoteProperty OnSubmission 0     -PassThru |
    Add-Member NoteProperty FirstSub     @()   -PassThru

$events = @()

function New-Event ($Date,$EventType,$WorkID) {
    New-Object psobject |
        Add-Member NoteProperty EventDate $Date      -PassThru |
        Add-Member NoteProperty EventType $EventType -PassThru |
        Add-Member NoteProperty WorkID    $WorkID    -PassThru
}

$submissions | foreach {
    $sub = $_
    $event = New-Event -Date $sub.DateSent -EventType "Submission" -WorkID $sub.WorkID
    $events += $event
    if ($sub.DateBack -match "\d\d\d\d\-\d\d\-\d\d") {
        if ($sub.DateBack -notlike "1899*") {
            if ($sub.Sale -eq 0) {
                $event = New-Event -Date $sub.DateBack -EventType "Rejection" -WorkID $sub.WorkID
            }
            else {
                $event = New-Event -Date $sub.DateBack -EventType "Sale" -WorkID $sub.WorkID
            }
            $events += $event
        }
    }
}

$stats = @()

$events | Sort-Object -Property EventDate | foreach {
    $event = $_
    if ($total.Date -eq $null) {
        $total.Date = $event.EventDate
    }
    #
    # should we output a new total record?
    if ($event.EventDate -ne $total.Date) {
        $total.Works = $worksHash.Count
        $total.FirstSub = $total.FirstSub -join ','
        $jsonString = $total | ConvertTo-Json
        $stats += $jsonString
        $total.Date = $event.EventDate
        $total.FirstSub = @()
    }
    switch ($event.EventType) {
        "Submission" {
                $total.OnSubmission += 1
            }
        "Rejection" {
                $total.OnSubmission -= 1
                $total.Rejections += 1
            }
        "Sale" {
                $total.OnSubmission -= 1
                $total.Acceptances += 1
            }
    }
    if ($worksHash[$event.WorkID] -eq $null) {
        $work = $works | where {$_.IDCode -eq $event.WorkID}
        $worksHash[$event.WorkID] = $work
        $total.FirstSub += ($work.Title -replace "%20",' ' -replace "%21",'!' -replace "%27",'''' -replace "%2f",'/')
    }
    $event | Out-Null
}
$total.Works = $worksHash.Count
$total.FirstSub = $total.FirstSub -join ','
$jsonString = $total | ConvertTo-Json
$stats += $jsonString

$bigJson = '[' + ($stats -join ',') + ']'

$alldata = $bigJson | ConvertFrom-Json

$alldata | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath "$sdir\SubmissionHistory.csv"

Write-Host "Done"