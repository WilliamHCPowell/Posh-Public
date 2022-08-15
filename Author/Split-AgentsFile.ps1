#
# script to re-purpose Agents.xlsx file as two sheets - Agencies and Agents
#

param ($JobName = "Local",
       $SubmissionData = '\\PHOENIX10\Bill\Documents\My Stories\Submissions\All Submissions.sonar3')

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
$AgenciesAgentsSubsFile       = "$sdir\AgenciesAgentsSubs.xlsx"
$NewSubmissionFile            = "$sdir\SubmissionReport2.xlsx"
#endregion

$Genres = "YA,MG,ScienceFiction,Fantasy,Historical,CrimeThriller,Upmarket,Literary,Romance,SupernaturalHorror" -split ','
$AgencyFields = "Agency,Sort,SonarID,Country,CoType,Pursue,Status,CompanyTwitter,ResponsePolicy,SubmissionURL,email,Notes" -split ','
$AgentFields = "Agency,AgentName,BestContact,Twitter,Pursue,Status,Tin2Gold,Thaumatist,Dom,BobbysDawn,ResponsePolicy,AgentURL,email,Exclusion,Notes" -split ','
$AgentFields = $AgentFields + $Genres
$AgentFieldsShort = $AgentFields | where {$_ -ne "BestContact"}

$Agencies = @{}
$Agents = @{}
$Works = @{}
$Submissions = @{}

$AgentsByMarketHash = @{}
$AllAgents = Import-Excel -Path $individualAgentsSheetFile
$AllAgents | foreach {
    $AgentRow = $_
    $AgencyRec = $AgentRow | Select-Object -Property $AgencyFields
    $Agencies[$AgencyRec.Agency] = $AgencyRec
    $AgentRec = $AgentRow | Select-Object -Property $AgentFields | foreach {$_.AgentName = $_.BestContact; $_} | Select-Object $AgentFieldsShort
    if (-not [string]::IsNullOrWhiteSpace($AgentRec.AgentName)) {
        $Agents[$AgentRec.AgentName] = $AgentRec
    }
<#
    if ([string]::IsNullOrWhiteSpace($AgentRow.SonarID)) {
        Write-Host "Agency '$($AgentRow.Agency)' does not have a Sonar ID"
    }
    else {
        $AgentRow.SonarID = [string]($AgentRow.SonarID)
        $AgentCollection = $AgentsByMarketHash[$AgentRow.SonarID]
        if ($AgentCollection -eq $null) {
            $AgentCollection = @($AgentRow)
        }
        else {
            $AgentCollection += $AgentRow
        }
        $AgentsByMarketHash[$AgentRow.SonarID] = $AgentCollection
    }
    $AgentRow | Out-Null
#>
}

if (Test-Path -LiteralPath $AgenciesAgentsSubsFile) {
    Remove-Item -LiteralPath $AgenciesAgentsSubsFile
}

$AgenciesFlat = $Agencies.GetEnumerator() | foreach {$_.Value} | Sort-Object -Property Sort

$WorksheetName = "Agencies and Publishers"

$AgenciesFlat | Export-Excel -Path $AgenciesAgentsSubsFile `
                                             -AutoSize `
                                             -AutoFilter `
                                             -WorksheetName $WorksheetName `
                                             -StartRow 1 `
                                             -StartColumn 1 `
                                             -FreezePane 2,8 `
                                             -TableName $($WorksheetName -replace "\s+",'' -replace "\p{P}+",'')

$AgentsFlat = $Agents.GetEnumerator() | foreach {$_.Value} | Sort-Object -Property Agency,AgentName

$WorksheetName = "Agents and Editors"

$AgentsFlat | Export-Excel -Path $AgenciesAgentsSubsFile `
                                             -AutoSize `
                                             -AutoFilter `
                                             -WorksheetName $WorksheetName `
                                             -StartRow 1 `
                                             -StartColumn 1 `
                                             -FreezePane 2,10 `
                                             -TableName $($WorksheetName -replace "\s+",'' -replace "\p{P}+",'')



Write-Host "Done"