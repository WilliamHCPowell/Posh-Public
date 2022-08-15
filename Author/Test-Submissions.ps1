#
# script to run daily to check when markets are open
#

param ($JobName = "Local",
       $SubmissionData = 'P:\Users\Bill\Documents\My Stories\Submissions\All Submissions.sonar3')

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
$attentionColours = @{foreground="black";background="green"}
$responseColours = @{foreground="black";background="yellow"}

#endregion


#region File Management
#
# Having identified the current working directory, we can now set up paths for the
# various cache files and report files used by the script.
#
# see https://blogs.msdn.microsoft.com/koteshb/2010/02/12/powershell-how-to-find-details-of-operating-system/
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
$MarketsToAddFile             = "$sdir\MarketsToAdd.csv"
$SubmissionScheduleFile       = "$sdir\SubmissionSchedule.json"
#endregion

$MarketsToAdd = @()

function New-AgentToAdd ($AgencyName,$SubmissionURL,$BestContact,$Guidelines) {
    New-Object PSObject |
        Add-Member NoteProperty AgencyName $AgencyName -PassThru |
        Add-Member NoteProperty SubmissionURL $SubmissionURL -PassThru |
        Add-Member NoteProperty BestContact $BestContact -PassThru |
        Add-Member NoteProperty Guidelines $Guidelines -PassThru
}

$AgentsByMarketHash = @{}
$AllAgents = Import-Excel -Path $individualAgentsSheetFile
$AllAgents | foreach {
    $AgentRow = $_
    if ([string]::IsNullOrWhiteSpace($AgentRow.SonarID)) {
        Write-Host "Agency '$($AgentRow.Agency)' does not have a Sonar ID"
        $Agent2add = New-AgentToAdd -AgencyName $AgentRow.Agency `
                                    -SubmissionURL $AgentRow.SubmissionURL `
                                    -BestContact $AgentRow.BestContact `
                                    -Guidelines "Agency,Novels"
        $MarketsToAdd += $Agent2add
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
}

$MarketCsvDataHash = @{}

if (Test-Path -LiteralPath $MarketFile) {
    $MarketCsvData = Import-Csv -LiteralPath $MarketFile
    $MarketCsvData | foreach {
        $rowData = $_
        if (-not [string]::IsNullOrWhiteSpace($rowData.SonarId)) {
            $MarketCsvDataHash[$rowData.SonarId] = $rowData
        }
        else {
            if ($rowData.Market -ne 'No') {
                Write-Host "Publisher $($rowData.Publisher) does not have a Sonar ID"
                $Agent2add = New-AgentToAdd -AgencyName $rowData.Publisher `
                                            -SubmissionURL $rowData.SubmitURL `
                                            -BestContact $rowData.BestContact `
                                            -Guidelines "Publisher,Novels"
                $MarketsToAdd += $Agent2add
            }
        }
    }
}

$MarketsToAdd | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $MarketsToAddFile

$text = Get-Content -LiteralPath $SubmissionData

$xmlDoc = [xml]$text

$TitlesToSkip = @(
        "Cricket",
        "Writers of the Future Anthology",
        "Interfiction Online",
        "Cemetery Dance",
        "Madeleine Milburn",
        "Georgina Capel Associates Ltd",
        "The Soho Agency",
        "The Ampersand Agency",
        "Terraform",
        "ASH Literary Agency",
        "Creative Authors"
)

$responsePolicyList = @()

$results = $xmlDoc.SONAR3.MARKETS.MARKET | foreach {
    $market = $_
    $title = [System.Web.HttpUtility]::UrlDecode($market.Title)
    if ($title -in @("Fantasy & Science Fiction (F&SF)","Ravenwood Quarterly","Constellary Tales")) {
        $title | Out-Null
    }
    $marketStatus = New-Object PSObject |
        Add-Member NoteProperty Title             $title           -PassThru |
        Add-Member NoteProperty SonarID           $market.IDCode   -PassThru |
        Add-Member NoteProperty IsAgency          $false           -PassThru |
        Add-Member NoteProperty IsDefunct         $false           -PassThru |
        Add-Member NoteProperty IsShort           $true            -PassThru |
        Add-Member NoteProperty URLList           @()              -PassThru |
        Add-Member NoteProperty ResponseContext   @()              -PassThru |
        Add-Member NoteProperty ResponsePolicy    ""               -PassThru |
        Add-Member NoteProperty SubmissionStatus  @()              -PassThru |
        Add-Member NoteProperty SentimentHash     @{}              -PassThru
    $Skip = $title -in $TitlesToSkip
    $MarketCsv = $null
    if (-not [string]::IsNullOrWhiteSpace($market.IDCode)) {
        $MarketCsv = $MarketCsvDataHash[$market.IDCode]
    }
    $AgentCollection = $AgentsByMarketHash[$market.IDCode]
    $Guidelines = $market.Guidelines -split ',' | foreach {
        $gl = $_
        switch ($gl) {
            "Agency" { $marketStatus.IsAgency = $true; $marketStatus.IsShort = $false }
            "Publisher" { $marketStatus.IsShort = $false }
            "Defunct" { $marketStatus.IsDefunct = $true }
        }
        $gl
    }
    if (<# ($marketStatus.IsShort) -and #> (-not $marketStatus.IsDefunct) -and (-not $Skip)) {
        #
        # try to identify the submission URL
        $allURLs = @{}
        "Address4,Address3,Address2,Address1,URL" -split ',' | foreach {
            $field = $_
            $url = $market.$field
            if ($url -like "*http*") {
                if ($url -like "*%*") {
                    $submitURL = [System.Web.HttpUtility]::UrlDecode($url)
                }
                else {
                    $submitURL = $url
                }
                $allURLs[$submitURL] = $true
            }
        }
        if ($MarketCsv -ne $null) {
            "TestOpenURL,GeneralURL,GuidelineURL,SubmitURL" -split ',' | foreach {
                $field = $_
                $url = $market.$field
                if ($url -like "*http*") {
                    if ($url -like "*%*") {
                        $submitURL = [System.Web.HttpUtility]::UrlDecode($url)
                    }
                    else {
                        $submitURL = $url
                    }
                    $allURLs[$submitURL] = $true
                }
            }
        }
        $AgentCollection | where {$_ -ne $null} | foreach {
            $AgentRow = $_
            "SubmissionURL,AgentURL" -split ',' | foreach {
                $field = $_
                $url = $AgentRow.$field
                if ($url -like "*http*") {
                    if ($url -like "*%*") {
                        $submitURL = [System.Web.HttpUtility]::UrlDecode($url)
                    }
                    else {
                        $submitURL = $url
                    }
                    $allURLs[$submitURL] = $true
                }
            }
        }
        if ($allURLs.Count -gt 0) {
            Write-Host @infoColours $title
            $allURLs.GetEnumerator() | foreach {
                $submitURL = $_.Key
                Write-Host @infoColours "Trying $submitURL"
                if ($submitURL -eq "https://cwagency.co.uk/page/submissions") {
                    $submitURL | Out-Null
                }
                try {
                 #   $info = Invoke-WebRequest -Uri $submitURL -TimeoutSec 20 -SessionVariable 'MySession' # -UseBasicParsing
                 #   $textToParse = $info.ParsedHtml.body.outerText
                    $info = Invoke-WebRequest -Uri $submitURL -TimeoutSec 20 -UseBasicParsing | ConvertFrom-Html
                    $textToParse = $info.InnerText
                    #
                    # look for submissions
                    $textToParse -split "[`r`n]+" -replace '&nbsp;',' ' -replace "^\s*",'' -replace "\s*$",'' -replace "\s+",' ' | where {$_ -like "*subm*"} | where {$_ -match "^.*(open|close|accept).*$"} | foreach {
                        $line = $_
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
                        elseif ($line -like "*no submission*") {
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
                    }
                    #
                    # look for response guidelines
                    $textToParse -split "[`r`n]+" -split "\." -replace '&nbsp;',' ' -replace "^\s*",'' -replace "\s*$",'' -replace "\s+",' ' | where {($_ -like "*respon*") -or 
                                                                                                                                                      ($_ -like "*hear from us*") -or 
                                                                                                                                                      ($_ -like "*be in touch*") -or 
                                                                                                                                                      ($_ -like "*aim to get back*")} | where {$_ -match "^*.(day|week|month).*$"} | foreach {
                        $line = $_
                        Write-Host @attentionColours $line
                        $responsePolicyList += $line
                        if ($line -match "^.*(if you have received no response|if you do not hear from us|if you don.t hear from us|if you have not heard|if we are interested|if we.re interested|if your submission is of interest).*$") {
                            $Cutoff = "Timeout:"
                        }
                        else {
                            $Cutoff = "Response:"
                        }
                        if ($line -match "^.*\b(?<periodLen>\d+|a|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|thirty|sixty) \b(?<periodType>day|days|week|weeks|month|months)\b.*$") {
                            if ([string]::IsNullOrWhiteSpace($marketStatus.ResponsePolicy)) {
                                $marketStatus.ResponsePolicy = $Cutoff + ' ' + $Matches["periodLen"] + ' ' + $Matches["periodType"]
                                $marketStatus.ResponseContext += $line
                                Write-Host @responseColours $marketStatus.ResponsePolicy
                            }
                        }
                    }

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
    $marketStatus
}

$results | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 -LiteralPath $MarketStatusFile

$responsePolicyList | foreach {
    $policyText = $_
    New-Object PSObject |
        Add-Member NoteProperty PolicyText $policyText -PassThru |
        Add-Member NoteProperty Relevant   ""          -PassThru |
        Add-Member NoteProperty IsTimeOut  ""          -PassThru |
        Add-Member NoteProperty Lower      ""          -PassThru |
        Add-Member NoteProperty Upper      ""          -PassThru |
        Add-Member NoteProperty TimeUnits  ""          -PassThru |
        Add-Member NoteProperty Spare      ""          -PassThru
} | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath "$sdir\ResponsePolicyList.csv"
Write-Host "Done"

exit 0

#
# See https://stackoverflow.com/questions/38005341/the-response-content-cannot-be-parsed-because-the-internet-explorer-engine-is-no
# Then run (from admin shell)
# Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -Value 2
#
$info = Invoke-WebRequest -Uri $URL # -UseBasicParsing 

$innerHtml = $info.ParsedHtml.body.innerHTML

