#
# script to get an up-to-date list of fiction markets from ralan.com
#

param ($SubmissionData = '\\PHOENIX10\Bill\Documents\My Stories\Submissions\All Submissions.sonar3',
       $RootURL="https://ralan.com/")

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
$RalanMarketsFile             = "$sdir\RalanMarkets.csv"
#endregion

$SonarIDLookup = @{}

if (Test-Path -LiteralPath $RalanMarketsFile) {
    Import-Csv -LiteralPath $RalanMarketsFile | where {[string]::IsNullOrWhiteSpace($_.SonarID)} | foreach {
        $RalanRec = $_
        $SonarIDLookup[$RalanRec.UniqueName] = $RalanRec
        $SonarIDLookup[$RalanRec.SonarID] = $RalanRec
    }
}

$marketPages = @(
    @{ href = "m.pro.htm"; Description = "Pro"; }, 
    @{ href = "m.semipro.htm"; Description = "Semipro"; }, 
    @{ href = "m.pay.htm"; Description = "Pay"; }, 
    @{ href = "m.token.htm"; Description = "Token"; }, 
    @{ href = "m.antho.htm"; Description = "Anthos"; }, 
    @{ href = "m.publish.htm"; Description = "Books"; }, 
    @{ href = "m.flash.htm"; Description = "Flash etc."; }, 
    @{ href = "m.contest.htm"; Description = "Contests"; }, 
    @{ href = "m.zzz.htm"; Description = "Sub-Static"; }
)

function Parse-Attributes ([string]$AttributeString) {
    #
    # clean up
    $AttributeString = $AttributeString -replace "^</A>",'' -replace " target=_blank",'' -replace "^\s*",'' -replace "\s*$",''
    # look for strings of the form \S+:\s*
    $fieldName = "Description"
    $stringSet = $AttributeString -split "\S+:\s*" -replace "^\s*",'' -replace "\s*$",''
    $offset = 0
    $stringBag = @{}
    $stringSet | foreach {
        $string = $_
        $stringBag[$fieldName] = $string
        $AttributeString = $AttributeString.Substring($string.Length, $AttributeString.Length - $string.Length) -replace "^\s*",'' -replace "\s*$",''
        $fieldNameLen = $AttributeString.IndexOf(':')
        if ($fieldNameLen -ne -1) {
            $fieldName = $AttributeString.Substring(0,$fieldNameLen) -replace "^\s*",'' -replace "\s*$",''
            $AttributeString = $AttributeString.Substring($fieldNameLen + 1, $AttributeString.Length - ($fieldNameLen + 1)) -replace "^\s*",'' -replace "\s*$",''
        }
    }
    $stringBag
}

$allAttributes = @{}

$AllMarkets = $marketPages | foreach {
    $marketPageInfo = $_
    $pageURL = "$RootURL/$($marketPageInfo.href)"
    $CategoryDescription = $marketPageInfo.Description

    $info = Invoke-WebRequest -Uri $pageURL # -UseBasicParsing 

    $innerHtml = $info.ParsedHtml.body.innerHTML
    $innerHtml -split '<A name=' | 
      where {$_ -like "*<LI>*"} |
    foreach {
        #
        # split into blocks according to name.
        $textBlock = $_
        $textBlock | Out-Null
        $markerEndOfUniqueName = $textBlock.IndexOf('>')
        $marketUniqueName = $textBlock.Substring(0,$markerEndOfUniqueName)
        $marketRest = $textBlock.Substring($markerEndOfUniqueName,$textBlock.Length - $markerEndOfUniqueName) -split "<LI>"
        $marketNews = $marketRest[0]
        $marketData = $marketRest[1] -split '<B>' -split '</B>'
        $marketURL = $marketData[0] -replace '<A href="','' -replace '" target=_blank>',''
        $marketTitle = $marketData[1]
        $marketAttributes = Parse-Attributes $marketData[2]
        $marketAttributes["MarketType"] = $CategoryDescription
        $marketAttributes["MarketTitle"] = $marketTitle
        $marketAttributes["UniqueName"] = $marketUniqueName
        $marketAttributes["URL"] = $marketURL
        $marketAttributes.GetEnumerator() | foreach {
            $allAttributes[$_.Key] = $true
        }
        $marketAttributes
    }
}

$marketFields = "MarketType,UniqueName,SonarID,MarketTitle,URL,Description" -split ','
$marketFields | foreach {
    $allAttributes[$_] = $false  # clear out the fields we already have in place
}

$allAttributes.GetEnumerator() | 
   foreach { if ($_.Value -eq $true) { $_.Key } } | 
   Sort-Object | 
   where {-not [string]::IsNullOrWhiteSpace($_)} |
   foreach { $marketFields += $_ }

$AllMarkets | foreach {
    $marketAttributes = $_
    $obj = New-Object PSObject | Select-Object -Property $marketFields
    $marketAttributes.GetEnumerator() | foreach {
        $obj.$($_.Key) = $_.Value
    }
    $obj
} | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $RalanMarketsFile

Write-Host "Done"
