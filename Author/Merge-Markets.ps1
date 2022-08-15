#
# script to merge market information between the Sonar submissions XML file and
# the Markets.csv
#

param ($SubmissionData = '\\PHOENIX8\Bill\Documents\My Stories\Submissions\All Submissions.sonar3')

[System.Reflection.Assembly]::LoadWithPartialName("System.web") | Out-Null

$infoColours = @{foreground = "cyan"}
$warningColours = @{foreground = "yellow"}
$errorColours = @{foreground = "red"}

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

#
# standard functions to find the directory in which the script is executing
# we'll use this info to read and write both cache files and reports
#
function Get-ScriptDirectory {
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
    Split-Path $Invocation.MyCommand.Path
}

$sdir = Get-ScriptDirectory
#endregion

#region File Management
#
# Having identified the current working directory, we can now set up paths for the
# various cache files and report files used by the script.
#
$SubmissionFile   = $sdir + "\" + "SubmissionReport.xlsx"              # These files are shared by many jobs
$MarketFile       = $sdir + "\" + "Markets.csv"                        #
$MarketUpdateFile = $sdir + "\" + "MarketsUpdate.csv"                  #

#endregion

$infoColours = @{foreground = "cyan"}
$warningColours = @{foreground = "yellow"}
$errorColours = @{foreground = "red"}

cls

$xmlText = Get-Content $SubmissionData
$xmlDoc = [xml]$xmlText

$marketData = Import-Csv -LiteralPath $MarketFile

$marketsBySonarId = @{}

$marketData | foreach {
    $marketRow = $_
    $marketsBySonarId[$marketRow.SonarID] = $marketRow
}

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


$xmlDoc.SONAR3.MARKETS.MARKET | foreach {
    $sonarMarket = $_
    $sonarTitle = [System.Web.HttpUtility]::UrlDecode($sonarMarket.Title)
    $sonarID = $sonarMarket.IDCode
    $marketRow = $marketsBySonarId[$sonarID]
    if ($marketRow -ne $null) {
        #
        # URL      = Home Page
        # Address1 = GuidelineURL
        # Address2 = GeneralURL
        # Address3 = TestOpenURL
        # Address4 = SubmitURL
        #
        Write-Host "Merge $($sonarTitle) and $($marketRow.Publisher)"
        #
        # test the URLs in the XML record first
        #
        "URL,Address1,Address2,Address3,Address4" -split ',' | foreach {
            $fieldName = $_
            $result = $null
            $URL = [System.Web.HttpUtility]::UrlDecode($sonarMarket.$fieldName)
            if (-not [string]::IsNullOrWhiteSpace($URL)) {
                Write-Host "Testing Field $fieldName : $URL"
                try {
                    $result = Invoke-WebRequest $URL -UseDefaultCredentials -TimeoutSec 15
                }
                catch {
                    $e = $_
                }
                if ($result -ne $null) {
                    if ($result.StatusCode -eq 200) {
                    }
                    else {
                        Write-Host "Bad status code $($result.StatusCode) from $($URL)"
                    }
                }
            }
        }
    }
    else {
        # Write-Host "No match for $($sonarTitle) in $MarketFile"
    }
}

Write-Host "Done"
