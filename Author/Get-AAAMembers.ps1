#
# script to get an up-to-date list of members of the Association of Authors' Agents
#

param ($URL="http://www.agentsassoc.co.uk/members-directory/")

cls

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
function Get-ScriptDirectory
{
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
# see https://blogs.msdn.microsoft.com/koteshb/2010/02/12/powershell-how-to-find-details-of-operating-system/
#
$associationMemberListFile    = "$sdir\AAAList.csv"
$individualAgentsListFile     = "$sdir\Agents.csv"
#endregion

#
# See https://stackoverflow.com/questions/38005341/the-response-content-cannot-be-parsed-because-the-internet-explorer-engine-is-no
# Then run (from admin shell)
# Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -Value 2
#
$info = Invoke-WebRequest -Uri $URL # -UseBasicParsing 

$innerHtml = $info.ParsedHtml.body.innerHTML

function Extract-URL ($RawString) {
    $agenturl = ""
    if ($RawString -match ".* href=""(https{0,1}://.*)"" target.*") {
        $agenturl = $Matches[1]
    }
    return $agenturl
}

$agentDetails = $innerHtml -split '<DIV class=shadow-box>' | Select-Object -skip 1 | foreach {
    $frag = $_
    $agentInfo = New-Object PSObject |
        Add-Member NoteProperty AgentName       ""     -PassThru |
        Add-Member NoteProperty AgentURL        ""     -PassThru |
        Add-Member NoteProperty AgentTwitter    ""     -PassThru |
        Add-Member NoteProperty AgentReferenced $false -PassThru

    $frag -split '<!-- ' | where {$_.Length -gt 0} | foreach {
        $element = [string]$_
        $elementTypeLen = $element.IndexOf(' -->')
        if ($elementTypeLen -ge 0) {
            $elementType = $element.Substring(0,$elementTypeLen)
            switch ($elementType) {
                "Member Logo" {
                     }
                "Member Name" {
                         if ($frag -match ".*<H2>(.*)</H2>.*") {
                             $agentInfo.AgentName = $Matches[1]
                             $agentInfo.AgentName = [System.Net.WebUtility]::HtmlDecode($agentInfo.AgentName)
                         }
                     }
                "Website" {
                         $agentInfo.AgentURL = Extract-URL $element
                     }
                "Twitter" {
                         $agentInfo.AgentTwitter = (Extract-URL $element) -replace 'https://twitter.com/',''
                     }
            }
        }
    }
    $agentInfo
}

$agentLookup = @{}

$agentDetails | foreach {
    $agentInfo = $_
    $agentLookup[$agentInfo.AgentName] = $agentInfo
}

$individualAgents = Import-Csv -LiteralPath $individualAgentsListFile

$individualAgents | foreach {
    $individualAgent = $_
    $agentInfo = $agentLookup[$individualAgent.Agency]
    if ($agentInfo -ne $null) {
        $individualAgent.IsAAA = $true
        $agentInfo.AgentReferenced = $true
    }
    else {
        $individualAgent.IsAAA = $false
    }
}

$individualAgents | 
    Sort-Object -Property @{Expression="Pursue";Descending=$true},@{Expression="Sort";Descending=$false} |
    Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $individualAgentsListFile

$agentDetails | 
    Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $associationMemberListFile

Write-Host "written $associationMemberListFile"
