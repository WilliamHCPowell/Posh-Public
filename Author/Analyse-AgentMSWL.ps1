#
# script to look up agent web-pages and analyse what they're looking for (MSWL)
#

param ($JobName = "Local",
       $SubmissionData = '\\PHOENIX10\Bill\Documents\My Stories\Submissions\All Submissions.sonar3')

cls

Add-Type -AssemblyName system.web

Import-Module ImportExcel
Import-Module PowerHTML

#
# see https://stackoverflow.com/questions/11696944/powershell-v3-invoke-webrequest-https-error
#
add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
$AllProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
[System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

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
$markerColours = @{foreground="magenta"}
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
$AgentDetailsFile             = "$sdir\AgentDetails.json"
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

$WorksheetName = "Agencies and Publishers"

$AgenciesFlat = Import-Excel -Path $AgenciesAgentsSubsFile -WorksheetName $WorksheetName

$WorksheetName = "Agents and Editors"

$AllAgents = Import-Excel -Path $AgenciesAgentsSubsFile -WorksheetName $WorksheetName

$AllAgentNames = $AllAgents.AgentName

#$AllAgents | ogv

$AgentInfo = $AllAgents | foreach {
    $Agent = $_
    $submitURL = $Agent.AgentURL
    if ($Agent.Pursue -eq 'Y') {
        if (-not [string]::IsNullOrWhiteSpace($submitURL)) {
                Write-Host @infoColours "Trying $submitURL"
                try {
                 #   $info = Invoke-WebRequest -Uri $submitURL -TimeoutSec 20 -SessionVariable 'MySession' # -UseBasicParsing
                 #   $textToParse = $info.ParsedHtml.body.outerText
                    $info = Invoke-WebRequest -Uri $submitURL -TimeoutSec 20 -UseBasicParsing | ConvertFrom-Html
                    $textToParse = $info.InnerText
                    $Twitter = ""
                    $likes = @()
                    $dislikes = @()
                    $allLines = $textToParse -split "[`r`n]+" -replace '&nbsp;',' ' -replace "^\s*",'' -replace "\s*$",'' -replace "\s+",' ' | where {-not [string]::IsNullOrWhiteSpace($_)} | foreach {
                        $line = [System.Web.HttpUtility]::HtmlDecode($_) -replace '\.com','ZZdotcomZZ'  -replace '\.co.uk','ZZdotcoukZZ' -replace 'e\.g\.','ZZegZZ' -replace '\.','. ' -replace '\.\s+','. ' -replace '\!','! ' -replace '\!\s+','! ' -replace '\?','? ' -replace '\?\s+','? '
                       # Write-Host $line
                        $line -split '[\!\?\.]\s+'
                    } | foreach {
                        $sentence = $_ -replace 'ZZegZZ','e.g.' -replace 'ZZdotcomZZ','.com' -replace 'ZZdotcoukZZ','.co.uk'
                        $sentence
                        if ($sentence -in $AllAgentNames) {
                            # put this in as a marker in both likes and dislikes
                            Write-Host @markerColours $sentence
                            $likes += $sentence
                            $dislikes += $sentence
                        }
                        elseif ($sentence -match ".*(not looking for|don.t represent|with the exception of).*") {
                            Write-Host @warningColours $sentence
                            $dislikes += $sentence
                        }
                        elseif ($sentence -match ".*(love|passion|like to read|personal interest|likes|relish|on the lookout for|is (also )*looking for|enjoy|keen|fortunate|represents|always open|can.t resist).*") {
                            Write-Host @debugColours $sentence
                            $likes += $sentence
                        }
                        elseif ($sentence -like "*I love*") {
                            Write-Host @debugColours $sentence
                            $likes += $sentence
                        }
                        elseif ($sentence -like "*I also love*") {
                            Write-Host @debugColours $sentence
                            $likes += $sentence
                        }
                        elseif ($sentence -like "*I also can't resist*") {
                            Write-Host @debugColours $sentence
                            $likes += $sentence
                        }
                        elseif ($sentence -like "*my list*") {
                            Write-Host @debugColours $sentence
                            $likes += $sentence
                        }
                        elseif ($sentence -like "*don't send*") {
                            Write-Host @warningColours $sentence
                            $dislikes += $sentence
                        }
<#
                        if ($line -like "*not accepting submissions from debut writers*") {
                            Write-Host @infoColours $line
                          #  $marketStatus.SubmissionStatus += $line
                          #  $marketStatus.SentimentHash["Closed"] = $true
                        }
                        elseif ($line -like "*not accepting*") {
                            Write-Host @errorColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Closed"] = $true
                        }
                        elseif ($line -like "*is closed*") {
                            Write-Host @errorColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Closed"] = $true
                        }
                        elseif ($line -like "*is now closed*") {
                            Write-Host @errorColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Closed"] = $true
                        }
                        elseif ($line -like "*are closed*") {
                            Write-Host @errorColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Closed"] = $true
                        }
                        elseif ($line -like "*are currently closed*") {
                            Write-Host @errorColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Closed"] = $true
                        }
                        elseif ($line -like "*are now closed*") {
                            Write-Host @errorColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Closed"] = $true
                        }
                        elseif ($line -like "*submissions closed*") {
                            Write-Host @errorColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Closed"] = $true
                        }
                        elseif ($line -like "*submission window closed*") {
                            Write-Host @errorColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Closed"] = $true
                        }
                        elseif ($line -like "*no open calls*") {
                            Write-Host @errorColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Closed"] = $true
                        }
                        elseif ($line -like "*currently closed*") {
                            Write-Host @errorColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Closed"] = $true
                        }
                        elseif ($line -like "*not currently open*") {
                            Write-Host @errorColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Closed"] = $true
                        }
                        elseif ($line -like "*is not open*") {
                            Write-Host @errorColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Closed"] = $true
                        }
                        elseif ($line -like "*not currently accepting*") {
                            Write-Host @errorColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Closed"] = $true
                        }
                        elseif ($line -like "*we accept submissions*") {
                            Write-Host @debugColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Open"] = $true
                        }
                        elseif ($line -like "*are open*") {
                            Write-Host @debugColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Open"] = $true
                        }
                        elseif ($line -like "*is open for*") {
                            Write-Host @debugColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Open"] = $true
                        }
                        elseif ($line -like "*is open submissions*") {
                            Write-Host @debugColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Open"] = $true
                        }
                        elseif ($line -like "*will be open to*") {
                            Write-Host @debugColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Future"] = $true
                        }
                        elseif ($line -like "*will be open for*") {
                            Write-Host @debugColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Future"] = $true
                        }
                        elseif ($line -like "*will be open from*") {
                            Write-Host @debugColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Future"] = $true
                        }
                        elseif ($line -like "*will open*") {
                            Write-Host @debugColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Future"] = $true
                        }
                        elseif ($line -like "*is currently open for*") {
                            Write-Host @debugColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Open"] = $true
                        }
                        elseif ($line -like "*is always open to*") {
                            Write-Host @debugColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Open"] = $true
                        }
                        elseif ($line -like "*open on*") {
                            Write-Host @debugColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Future"] = $true
                        }
                        elseif ($line -like "*open to*") {
                            Write-Host @debugColours $line
                            $marketStatus.SubmissionStatus += $line
                            $marketStatus.SentimentHash["Open"] = $true
                        }
                        else {
                           # Write-Host @infoColours $line
                        }
#>
                    }
                    New-Object PSObject |
                        Add-Member NoteProperty Agent     $Agent.AgentName -PassThru |
                        Add-Member NoteProperty Agency    $Agent.Agency    -PassThru |
                        Add-Member NoteProperty Twitter   $Agent.Twitter   -PassThru |
                        Add-Member NoteProperty Likes     $likes           -PassThru |
                        Add-Member NoteProperty Dislikes  $dislikes        -PassThru |
                        Add-Member NoteProperty AgentText $allLines        -PassThru

                }
                catch {
                    $e = $_
                    $exceptionText = $e.Exception.Message
                    Write-Host @warningColours "$submitURL : $exceptionText"
                    $marketStatus.URLList += New-Object PSObject |
                                                 Add-Member NoteProperty URL              $submitURL     -PassThru |
                                                 Add-Member NoteProperty ExceptionMessage $exceptionText -PassThru
                }
        }
    }
}

#
# see https://www.cryingcloud.com/blog/2017/05/02/replacefix-unicode-characters-created-by-convertto-json-in-powershell-for-arm-templates
$AgentInfo | ConvertTo-Json -Depth 6  | % { [System.Text.RegularExpressions.Regex]::Unescape($_) } | Set-Content -Encoding UTF8 -LiteralPath $AgentDetailsFile

Write-Host "Done"
