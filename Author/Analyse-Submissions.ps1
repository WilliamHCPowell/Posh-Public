#
# script to analyse submissions to various markets
#
# DONE - calculate which magazines a story can go to based on wordlength
# TODO - check for missing URLs and guidelines in markets
# DONE - mark as Blocked magazines which have a story under submission
# TODO - count per story of how many submissions
# TODO - count per market of how many submissions
# PART - Freeze Panes, shading
# PART - eliminate horror / quiltbag other mismatches
# DONE - Test for the existence of submission pages & look for the word "CLOSED"
#

param ($SubmissionData = '\\PHOENIX10\Bill\Documents\My Stories\Submissions\All Submissions.sonar3',
    [switch]$HideIneligible,
    [switch]$HideOut,
    [switch]$NoUrlTest,
    $GlobalFlipFlag = $true, # when true flips rows and columns so markets are on left and works along top,
    [switch]$runTests, # just runs URL tests
    [switch]$UseInvokeWR,
    [switch]$SlowCompare,
    $JobName = "Local"
)

Add-Type -AssemblyName system.web

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

#endregion

#region Excel Constants
#
# Chart types, etc, defined at http://it.toolbox.com/wiki/index.php/EXCEL_Chart_Type_Enumeration
$xlLine = 4
$xlBarClustered = 57
$xlColumnClustered = 51
$xl3DColumnStacked = 55
$xl3DColumn = -4100

$xlRows = 1
$xlColumns = 2

$xlLocationAsObject = 2

$xlCategory = 1
$xlPrimary = 1
$xlValue = 2
$xlSeriesAxis = 3

$xlLocationAsNewSheet = 1
$xlRight = -4152
$xlBuiltIn = 21

$xlTickMarkCross = 4    # Crosses the axis.
$xlTickMarkInside = 2   # Inside the axis.
$xlTickMarkNone = -4142 # No mark.
$xlTickMarkOutside = 3  # Outside the axis.

$xlScaleLogarithmic = -4133
$xlScaleLinear = -4132

$xlMaximized = -4137
$xlMinimized = -4140
$xlNormal = -4143

$xlA1 = 1
$xlR1C1 = -4150
#endregion

$today = date

$SonarIDLookup = @{}

if (Test-Path -LiteralPath $RalanMarketsFile) {
    Import-Csv -LiteralPath $RalanMarketsFile | where {-not [string]::IsNullOrWhiteSpace($_.SonarID)} | foreach {
        $RalanRec = $_
        $SonarIDLookup[$RalanRec.UniqueName] = $RalanRec
        $SonarIDLookup[$RalanRec.SonarID] = $RalanRec
    }
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
            Write-Host "Publisher $($rowData.Publisher) does not have a Sonar ID"
        }
    }
}

$AgencyInfoHash = @{}

if (Test-Path -LiteralPath $AgentInfoFile) {
    Import-Csv -LiteralPath $AgentInfoFile | foreach {
        $agentInfo = $_
        if (-not [string]::IsNullOrWhiteSpace($agentInfo.SonarID)) {
            $agencyInfo = $AgencyInfoHash[$agentInfo.SonarID]
            if ($agencyInfo -eq $null) {
                $agencyInfo = @($agentInfo)
            }
            else {
                $agencyInfo += $agentInfo
            }
            $AgencyInfoHash[$agentInfo.SonarID] = $agencyInfo
        }
    }
}

#region MarketStatus

$now = Get-Date
$sevenDaysAgo = $now.AddDays(-7)

$MarketStatusFileItem = Get-Item -LiteralPath $MarketStatusFile
if ($MarketStatusFileItem.LastWriteTime -lt $sevenDaysAgo) {
    .\Test-Submissions.ps1
}

$MarketStatusCollection = Get-Content -LiteralPath $MarketStatusFile | ConvertFrom-Json

$MarketStatusLookup = @{}
$MarketStatusCollection | foreach {
    $MarketStatusRow = $_
    $MarketStatusLookup[$MarketStatusRow.SonarID] = $MarketStatusRow
}

#endregion

$SubXml = Get-Content -LiteralPath $SubmissionData

$doc = [xml]$SubXml

$root = $doc.SONAR3

$ineligibleHash = @{}
$outHash = @{}
$marketSynthByTitle = @{}
$marketSynthByIDCode = @{}
$marketIndexByIDCode = @{}

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

#
# See https://gist.github.com/angel-vladov/9482676
function Read-HtmlPage {
    param ([Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)][String] $Uri)

    # Invoke-WebRequest and Invoke-RestMethod can't work properly with UTF-8 Response so we need to do things this way.
    [Net.HttpWebRequest]$WebRequest = [Net.WebRequest]::Create($Uri)
    $WebRequest.Timeout = 15 * 1000  # milliseconds
    [Net.HttpWebResponse]$WebResponse = $WebRequest.GetResponse()
    $Reader = New-Object IO.StreamReader($WebResponse.GetResponseStream())
    $Response = $Reader.ReadToEnd()
    $Reader.Close()

    # Create the document class
    [mshtml.HTMLDocumentClass] $Doc = New-Object -com "HTMLFILE"
    $Doc.IHTMLDocument2_write($Response)
    
    # Returns a HTMLDocumentClass instance just like Invoke-WebRequest ParsedHtml
    $Doc
}

function Test-URL-Invoke-WebRequest ([string]$URL) {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12;
    $page = Invoke-WebRequest $URL -UseDefaultCredentials -TimeoutSec 15
    return $page.ParsedHtml.body.outerText
}

function Test-URL-Read-HtmlPage ([string]$URL) {
    $page = Read-HtmlPage -Uri $URL
    $unparsedText = $page.all | foreach { $_.outerText }
    $unparsedText = $unparsedText -join "`n"
    return $unparsedText
}

$connectionErrors = @{}

function Test-URL ([string]$URL, [bool]$UseInvokeWR = $false) {
    $URL = $URL -replace "^\s*", '' -replace "\s*$", ''
    if ([string]::IsNullOrWhiteSpace($URL)) {
        return "Blank"  
    }
    elseif ($URL.StartsWith("mailto:")) {
        return "Mail"  
    }
    else {
        $retVal = "Open"
    }
    $domain = $URL -replace 'https://','' -replace 'http://','' -replace "\/.*$",''
    try {
        $domainInfo = Resolve-DnsName $domain -ErrorAction SilentlyContinue
        if ($domainInfo -eq $null) {
            return "No Domain"
        }
    }
    catch {
        $e = $_
        return "No Domain"
    }


    try {
        if ($UseInvokeWR) {
            $textToCheck = Test-URL-Invoke-WebRequest -URL $URL
        }
        else {
            $textToCheck = Test-URL-Read-HtmlPage -URL $URL
        }
    }
    catch {
        $e = $_
        Write-Host @errorColours "Error : $e"
        Write-Host @errorColours "URL : $URL"
        switch -Wildcard ($e) {
            "*404*" {
                $retVal = "Not Found"
                break
            }
            "*400*" {
                $retVal = "Not Found"
                break
            }
            "*Not Found*" {
                $retVal = "Not Found"
                break
            }
            "*could not be resolved*" {
                $retVal = "Not Found"
                break
            }
            "*503*" {
                $retVal = "Busy"
                break
            }
            "*Could not create SSL/TLS*" {
                $retVal = "SSL"
                break
            }
            default {
                $retVal = "Fault"
                break
            }
        }
        $connectionErrors[$URL] = $retval
    }

    if ($textToCheck -ne $null) {
        #
        # for the sake of brevity, limit the length of text to be checked
        $textToCheck = $textToCheck.Substring(0, [math]::Min($textToCheck.Length, 4096))
        switch -regex ($textToCheck) {
            ".*Closed.*" {
                $retVal = "Closed"
            }
            ".*not accepting.*" {
                $retVal = "Closed"
            }
            ".*no open calls.*" {
                $retVal = "Closed"
            }
            ".*on hiatus.*" {
                $retVal = "Closed"
            }
        }
    }


    $retVal
}

if ($runTests) {
    Test-URL "http://mothershipzeta.org/submission-guidelines/" -UseInvokeWR: $UseInvokeWR
    Test-URL "http://mothershipzeta.org/submission-guidelines/index.zzz" -UseInvokeWR: $UseInvokeWR
    Test-URL "http://steampunkuniverse.alliterationink.com/" -UseInvokeWR: $UseInvokeWR
    Test-URL "http://aresmagazine.com/call-for-content/submission-guidelines/" -UseInvokeWR: $UseInvokeWR
    Test-URL "https://www.podomatic.com/podcasts/beammeup" -UseInvokeWR: $UseInvokeWR
    exit 0
}

$termHash = @{}

function Analyse-Guidelines ([string]$Guidelines) {
    $obj = New-Object PSObject
    $gls = $Guidelines -split ","
    $gls = @($gls)
    foreach ($gl in $gls) {
        #    if ($gl -eq "Audio") {
        #      Write-Host $gl
        #    }
        $termHash[$gl]++
        switch -regex ($gl) {
            "(Publisher)|(Agency)|(Anthology)|(Audio)|(Themed)" {
                $obj = $obj | Add-Member NoteProperty  MarketType  $gl -Force -PassThru
            }
            "(Closed)|(Unreachable)|(Defunct)|(Ineligible)|(Hiatus)" {
                $obj = $obj | Add-Member NoteProperty  Closed      $gl -Force -PassThru
            }
            "(SFWA)|(Pro)|(Semi)|(Unpaid)|(Royalty)|(Token)|(Flat)|(\d+c)" {
                $obj = $obj | Add-Member NoteProperty  Rate        $gl -Force -PassThru
            }
            "Open=.*" {
                $obj = $obj | Add-Member NoteProperty  OpenDate    $gl -Force -PassThru
            }
            "Deadline=.*" {
                $obj = $obj | Add-Member NoteProperty  Deadline    $gl -Force -PassThru
                break # so we don't process it as a range of numbers also
            }
            "Target=.*" {
                $obj = $obj | Add-Member NoteProperty  Target      $gl -Force -PassThru
            }
            "^((Horror)|(Steampunk)|(SF)|(Fantasy)|(YA)|(MG)|(Detective)|(Humour)|(SFF)|(Alt History)|(Detective)|(Mystery)|(Weird)|(Paranormal)|(Historical))$" {
                $obj = $obj | Add-Member NoteProperty  Genre       $gl -Force -PassThru
            }
            "Parallel" {
                $obj = $obj | Add-Member NoteProperty  Parallel    $gl -Force -PassThru
            }
            "QUILTBAG" {
                $obj = $obj | Add-Member NoteProperty  Affirmative $gl -Force -PassThru
            }
            "([\d]+\-[\d]+)" {
                $obj = $obj | Add-Member NoteProperty  WordRange   $gl -Force -PassThru
            }
            "([\d]+\+)" {
                $obj = $obj | Add-Member NoteProperty  WordRange   $gl -Force -PassThru
            }
            "(Short)|(Flash)|(Novels*)" {
                $obj = $obj | Add-Member NoteProperty  Category    $gl -Force -PassThru
            }
            "\S+\s+\S+" {
                $obj = $obj | Add-Member NoteProperty  Tagline     $gl -Force -PassThru
            }
            "^\s*$" {
                # silently ignore empty guidelines
            }
            default {
                if (-not ([string]$gl -match "^\s*$")) {
                    $obj = $obj | Add-Member NoteProperty  Other       $gl -Force -PassThru
                }
            }
        }
    }
    $obj | Select-Object -Property MarketType, Closed, Rate, OpenDate, Deadline, Target, Genre, Parallel, Affirmative, WordRange, Category, Tagline, Other
}

function Split-WordRange ([string]$WordRange) {
    $bounds = $WordRange -split "[\+\-]" | foreach {
        if ([string]$_ -eq "") {
            [int]::MaxValue
        }
        else {
            [int]$_
        }
    }
    return $bounds
}

$marketsSynthUnsorted = $root.MARKETS.MARKET | foreach {
    $m = $_
    #  if ($m.IDCode -eq 69512) {
    #    Write-Host "Hoo!"
    #  }
    $guidelinesText = [System.Web.HttpUtility]::UrlDecode($m.Guidelines)
    $g = Analyse-Guidelines -Guidelines $guidelinesText
    $o = New-Object PSObject | 
        Add-Member NoteProperty Type  "Market"     -PassThru | 
        Add-Member NoteProperty Index $marketIndex -PassThru
    'IDCode,Title,Editor,Email,Phone,URL,Fax,Comments,Guidelines,LastUpdated,Address1,Address2,Address3,Address4' -split ',' | foreach {
        $pn = $_
        $prop = [System.Web.HttpUtility]::UrlDecode($m."$pn")
        $o = $o |
            Add-Member NoteProperty -Name $pn -Value $prop -PassThru
    }
    'MarketType,Closed,Rate,OpenDate,Deadline,Target,Genre,Parallel,Affirmative,WordRange,Category,Tagline,Other' -split ',' | foreach {
        $pn = $_
        if ($pn -eq "WordRange") {
            $prop = $g."$pn"
        }
        else {
            $prop = [System.Web.HttpUtility]::UrlDecode($g."$pn")
        }
        $o = $o |
            Add-Member NoteProperty -Name $pn -Value $prop -PassThru
    }
    $o
} | foreach {
    #
    # add aliases for Address1-4
    $_ |
        Add-Member AliasProperty -Name GuidelineURL -Value Address1  -PassThru |
        Add-Member AliasProperty -Name GeneralURL   -Value Address2  -PassThru |
        Add-Member AliasProperty -Name TestOpenURL  -Value Address3  -PassThru |
        Add-Member AliasProperty -Name SubmitURL    -Value Address4  -PassThru |
        Add-Member AliasProperty -Name WordSpan     -Value WordRange -PassThru |
        Add-Member AliasProperty -Name Addressline1 -Value Address1  -PassThru |
        Add-Member AliasProperty -Name Addressline2 -Value Address4  -PassThru |
        Add-Member AliasProperty -Name Addressline3 -Value Address3  -PassThru |
        Add-Member AliasProperty -Name Addressline4 -Value Address2  -PassThru
}

$categorySort = @{}
$categorySort["Flash"] = 0
$categorySort["Short"] = 1
$categorySort["Novels"] = 2

$rateSort = @{}
$rateSort["SFWA"] = 0
$rateSort[""] = 1
$rateSort["Flat"] = 2
$rateSort["Royalty"] = 3
$rateSort["Unpaid"] = 4

$CopyToXml = 0
$CopyToCsv = 1
$DontCopy = 2
$QueryCopy = 3

#
# note that on the form the Address fields are in the order 1,4,3,2
$fieldMap = @(
    # Xml            Csv             CopyDirection
    @("Title", "Publisher", $DontCopy),
    @("URL", "GeneralURL", $CopyToXml),
    @("Address1", "GuidelineURL", $CopyToXml),
    @("Address2", "GeneralURL", $CopyToXml),
    @("Address3", "TestOpenURL", $CopyToXml),
    @("Address4", "SubmitURL", $CopyToXml),
    @("WordRange", "WordRange", $CopyToXml),
    @("Genre", "Genre", $QueryCopy),
    @("Tagline", "Tagline", $QueryCopy)
    @("Category", "Category", $CopyToCsv)
)

$copy2xml = @{foreground = "cyan"}
$copy2csv = @{foreground = "magenta"}

$noMarketData = @{}

function Correlate-Data ($MarketXml) {
    $count = 0
    $MarketCsv = $MarketCsvDataHash[$MarketXml.IDCode]
    if ($MarketCsv -eq $null) {
        Write-Host "No market data found for $($MarketXml.Title), ID $($MarketXml.IDCode)" 
        $noMarketData[$MarketXml.IDCode] = $MarketXml.Title
        return
    }
    #
    # manage equivalent fields
    $fieldMap | foreach {
        $mapLine = $_
        $xmlField = $mapLine[0]
        $csvField = $mapLine[1]
        $xmlVal = [string]($MarketXml."$xmlField")
        $csvVal = [string]($MarketCsv."$csvField")
        if (($xmlVal -eq "") -and ($csvVal -ne "")) {
            Write-Host @copy2xml "Overwrite for $($MarketXml.Title): Xml.$xmlField = $xmlVal <= Csv.$csvField = $csvVal"
            $MarketXml."$xmlField" = $csvVal
            $count++
        }
        elseif (($xmlVal -ne "") -and ($csvVal -eq "")) {
            Write-Host @copy2csv "Overwrite for $($MarketXml.Title): Xml.$xmlField = $xmlVal => Csv.$csvField = $csvVal"
            $MarketCsv."$csvField" = $xmlVal
            $count++
        }
        elseif ($xmlVal -ne $csvVal) {
            switch ($mapLine[2]) {
                $CopyToXml {
                    Write-Host @copy2xml "Mismatch for $($MarketXml.Title): Xml.$xmlField = $xmlVal <= Csv.$csvField = $csvVal"
                    $MarketXml."$xmlField" = $csvVal
                    $count++
                }
                $CopyToCsv {
                    Write-Host @copy2csv "Mismatch for $($MarketXml.Title): Xml.$xmlField = $xmlVal => Csv.$csvField = $csvVal"
                    $MarketCsv."$csvField" = $xmlVal
                    $count++
                }
                $DontCopy {
                    Write-Host "Ignore Mismatch for $($MarketXml.Title): Xml.$xmlField = $xmlVal ; Csv.$csvField = $csvVal"
                }
                $QueryCopy {
                    Write-Host "Mismatch for $($MarketXml.Title): Xml.$xmlField = $xmlVal ; Csv.$csvField = $csvVal"
                    $x = Read-Host "Continue"
                }

            }
        } # else they must be the same, whether blank or non-blank - ignore
    }
    #
    # manage derived fields
    switch ($MarketXml.Rate) {
        "SFWA" {
            if ([string]$MarketCsv.SFWA -eq "") {
                $MarketCsv.SFWA = "Y"
            }
        }
    }
    #
    # if there's no rate, see if we can get it from the market csv
    if ([string]::IsNullOrWhiteSpace($MarketXml.Rate)) {
        if (-not [string]::IsNullOrWhiteSpace($MarketCsv.Rate)) {
            # Write-Host "Set Rate"
            switch -Wildcard ($MarketCsv.PaymentType) {
               "*flat*" {
                       $MarketXml.Rate = "Flat"
                       $count++
                   }
               "*unpaid*" {
                       $MarketXml.Rate = "Unpaid"
                       $count++
                   }
               "cents (US) per word" {
                       $MarketXml.Rate = $MarketCsv.Rate + 'c'
                       $count++
                   }
               "cents (CAD) per word" {
                       $MarketXml.Rate = $MarketCsv.Rate + 'c (CAD)'
                       $count++
                   }
               "cents (AUD) per word" {
                       $MarketXml.Rate = $MarketCsv.Rate + 'c (AUD)'
                       $count++
                   }
               default {
                       if ($MarketCsv.SFWA -eq "Y") {
                           $MarketXml.Rate = "SFWA"
                           $count++
                       }
                   }
            }
        }
    }
    if ([string]$MarketXml.WordRange -ne "") {
        $bounds = Split-WordRange $MarketXml.WordRange
        if ($bounds.GetType().Name -eq "Int32") {
            Write-Host "error on split: $($MarketXml.Title): $($MarketXml.WordRange)"
        }
        else {
            $MarketCsv.MinWords = $bounds[0]
            $MarketCsv.MaxWords = $bounds[1]
        }
    }
    $count
}

#region Market
#
# split out the guidelines for each market
#
$marketIndex = 0
$marketsSynthSorted = $marketsSynthUnsorted | Sort-Object -Property @{Expression = {[int]$categorySort[($_.Category)]}; Descending = $false}, MarketType, @{Expression = {[int]$rateSort[[string]($_.Rate)]}; Descending = $false}, Title | foreach {
    $market = $_
    #  if ($market.IDCode -eq 69512) {
    #    Write-Host "Hoo!"
    #  }
    #
    # correlate with the market data from the csv file
    $changes = Correlate-Data -MarketXml $market
    $market.Index = $marketIndex
    $marketSynthByTitle[$market.Title] = $market
    $marketSynthByIDCode[$market.IDCode] = $market
    $marketIndexByIDCode[$market.IDCode] = $marketIndex++
    if (($changes -gt 0) -and $SlowCompare) {
        $market | Out-Host
        $x = Read-Host "Press return to continue"
    }
    $market
}

$marketsSynthSorted | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath "$sdir\SonarMarkets.csv"

#
# re-sort by Category, then title
#$marketIndex = 0
#$marketsSynthSorted = $marketsSynthSorted | Sort-Object -Property SortCategory,Title | foreach {
#  $m = $_
#  $m.Index = $marketIndex++
#  $m
#}
#endregion

$marketListUnsorted = foreach ($item in $MarketCsvDataHash.GetEnumerator()) {
    $item.Value
}

$marketListUnsorted | Sort-Object -Property Publisher | Export-Csv -NoTypeInformation -Encoding UTF8 $MarketUpdateFile

$doc.Save("NewSonar.Xml")

if ($SlowCompare) {
    exit 0
}
#####################################################################################################
#####################################################################################################
#####################################################################################################
#####################################################################################################
#####################################################################################################

#region Works
#
# Loop through the works
#
$workHash = @{}
$workRow = @{}

$workIndex = 0
$works = $root.WORKS.WORK | Sort-Object -Property @{Expression = {[int]($_.Words)}; Descending = $false} | foreach {
    $work = $_
    if ([int]$work.Words -lt 1000) {
        $category = "Flash"
    }
    elseif ([int]$work.Words -le 40000) {
        $category = "Short"
    }
    else {
        $category = "Novels"
    }
    $work2 = $work.Clone() | 
        Add-Member NoteProperty Type      "Work"     -PassThru | 
        Add-Member NoteProperty Category  $category  -PassThru | 
        Add-Member NoteProperty Index     $workIndex -PassThru
    "Title", "Description", "Genre", "Filename", "Comments" | foreach {
        $Field = $_
        $work2."$Field" = [System.Web.HttpUtility]::UrlDecode($work2.$Field)
    }
    #  $work2.Words = [int]$work2.Words    # force integer
    $workHash[$work2.IDCode] = $work2
    $workRow[$work2.IDCode] = $workIndex++
    $work2
}
#endregion

#region Utilities
function Set-CellValue ($Row, $Col, $Value) {
    # optimise empty string writes...
    if ([string]$Value -match "^\s*$") {
        return
    }
    if ($GlobalFlipFlag) {
        $worksheet.Cells.Item($Col, $Row) = $Value
    }
    else {
        try {
            $worksheet.Cells.Item($Row, $Col) = $Value
        }
        catch {
            $e = $_
            Write-Host "Error : $e"
        }
    }
}

function Set-CellProperty ($Row, $Col, $Property, $Value) {
    $props = $Property -split "\."
    $props = @($props)
    switch ($props.Length) {
        1 {
            if ($GlobalFlipFlag) {
                $worksheet.Cells.Item($Col, $Row).$($props[0]) = $Value
            }
            else {
                $worksheet.Cells.Item($Row, $Col).$($props[0]) = $Value
            }
        }
        2 {
            if ($GlobalFlipFlag) {
                $worksheet.Cells.Item($Col, $Row).$($props[0]).$($props[1]) = $Value
            }
            else {
                $worksheet.Cells.Item($Row, $Col).$($props[0]).$($props[1]) = $Value
            }
        }
        default {
            Write-Host "$Property not supported"
        }
    }
}

function Set-CellMethodCall ($Row, $Col, $Method, $Value, [switch]$NoValue) {
    if ($NoValue) {
        if ($GlobalFlipFlag) {
            $worksheet.Cells.Item($Col, $Row).$Method() | out-null
        }
        else {
            try {
                $worksheet.Cells.Item($Row, $Col).$Method() | out-null
            }
            catch {
                $e = $_
                Write-Host "Error : $e"
            }
        }
    }
    else {
        if ($GlobalFlipFlag) {
            $worksheet.Cells.Item($Col, $Row).$Method($Value) | out-null
        }
        else {
            $worksheet.Cells.Item($Row, $Col).$Method($Value) | out-null
        }
    }
}

function Set-ColumnProperty ($Col, $Property, $Value) {
    $props = $Property -split "\."
    $props = @($props)
    switch ($props.Length) {
        1 {
            $worksheet.columns.Item($Col).$($props[0]) = $Value
        }
        2 {
            $worksheet.columns.Item($Col).$($props[0]).$($props[1]) = $Value
        }
        default {
            Write-Host "$Property not supported"
        }
    }
}

function Set-RowProperty ($Row, $Property, $Value) {
    $props = $Property -split "\."
    $props = @($props)
    switch ($props.Length) {
        1 {
            $worksheet.rows.Item($Row).$($props[0]) = $Value
        }
        2 {
            $worksheet.rows.Item($Row).$($props[0]).$($props[1]) = $Value
        }
        default {
            Write-Host "$Property not supported"
        }
    }
}

#endregion

Out-File ".\${JobName}_Pre-Excel_Done.timestamp"

#region Submissions-Setup
#
# save all the submissions
#
$submissionHash = @{}

if (Test-Path -LiteralPath $SubmissionFile) {
    Remove-Item -LiteralPath $SubmissionFile
}
$Excel = New-Object -ComObject excel.application

$IsNonInteractiveWindow = [bool]([Environment]::GetCommandLineArgs() -like '-noni*')
if ($IsNonInteractiveWindow) {
    $Excel.visible = $false
}
else {
    $Excel.visible = $true
}
$workbook = $Excel.Workbooks.Add() 
#
# maximise the window so that freeze panes works
$win = $Excel.ActiveWindow
$win.WindowState = $xlMaximized

$workSheetList = @(
    @{Name = "Short Stories - Pro";   Index = 1; MinWordCount  = 0 ;     MaxWordCount = 40000 ;  IsPro = $true;  IsNovel = $false ; Scale = 75 ; MarketList = $null; },
    @{Name = "Short Stories - Other"; Index = 2; MinWordCount  = 0 ;     MaxWordCount = 40000 ;  IsPro = $false; IsNovel = $false ; Scale = 75 ; MarketList = $null; },
    @{Name = "Novels";                Index = 3; MinWordCount  = 40001 ; MaxWordCount = 200000 ; IsPro = $null;  IsNovel = $true ;  Scale = 75 ; MarketList = $null; },
    @{Name = "Defunct";               Index = 4; MinWordCount  = 0 ;     MaxWordCount = 40000 ;  IsPro = $null;  IsNovel = $true ;  Scale = 75 ; MarketList = $null; }
)

while ($workbook.Sheets.Count -lt $workSheetList.Length) {
    $workbook.worksheets.Add() | out-null
}

$app = $workbook.Application

#
# determine the geometry of the spreadsheet
#
$RowIndex = 1
$TitleRow = $RowIndex++
$IDRow = $RowIndex++
$URLRow = $RowIndex++
$URLStateRow = $RowIndex++
#$URLSubmissionRow = $RowIndex++
#$URLGuidelineRow = $RowIndex++
#$URLGeneralRow = $RowIndex++
$BlockedRow = $RowIndex++
$MismatchRow = $RowIndex++
$ClosedRow = $RowIndex++
$OpenRow = $RowIndex++
$DeadlineRow = $RowIndex++
$RateRow = $RowIndex++
$WordsRow = $RowIndex++
$TypeRow = $RowIndex++
$CategoryRow = $RowIndex++
$GenreRow = $RowIndex++
$TaglineRow = $RowIndex++
$AffirmativeRow = $RowIndex++
$ParallelRow = $RowIndex++
$OtherRow = $RowIndex++
$GuideLinesRow = $RowIndex++
$HeaderRow = $RowIndex++
$BaseRow = $RowIndex

$worksheetIndex = 0

$novelMarkets = @()
$proMarkets = @()
$otherMarkets = @()
$defunctMarkets = @()

$marketsSynthSorted | foreach {
    $m = $_
    if ($m.Category -eq "Novels") {
        $novelMarkets += $m
    }
    elseif ($m.Closed -eq "Defunct") {
        $defunctMarkets += $m
    }
    else {
        $RalanRec = $SonarIDLookup[$m.IDCode]
        if ($m.Rate -eq "SFWA") {
            $proMarkets += $m
        }
        elseif ($RalanRec -eq $null) {
            $otherMarkets += $m
        }
        elseif ($RalanRec.MarketType -eq "Pro") {
            $proMarkets += $m
        }
        else {
            $otherMarkets += $m
        }
    }
}

$workSheetList | foreach {
    $worksheetInfo = $_
    $worksheetIndex++

    #
    # subset the market, work and submission data
    $marketIndex = 0
    $marketIndexByIDCode = @{}
    switch ($worksheetInfo.Name) {
        "Short Stories - Pro" { $marketsSynthSorted = $proMarkets }
        "Novels" { $marketsSynthSorted = $novelMarkets }
        "Defunct" { $marketsSynthSorted = $defunctMarkets }
        default { $marketsSynthSorted = $otherMarkets }
    }
    $marketSubset = $marketsSynthSorted | foreach {
        $m = $_
        $m.Index = $marketIndex
        $marketIndexByIDCode[$m.IDCode] = $marketIndex++
        $m
    }

    $workIndex = 0
    $workRow = @{}
    $workSubset = $works | 
        where {[int]$_.Words -ge $worksheetInfo.MinWordCount} | 
        where {[int]$_.Words -le $worksheetInfo.MaxWordCount} | 
        foreach {
        $w = $_
        $w.Index = $workIndex
        $workRow[$w.IDCode] = $workIndex++
        $w
    }

    $submissionSubset = $root.SUBMISSIONS.SUBMISSION | foreach {
        $sub = $_
        if ($marketIndexByIDCode[$sub.MarketID] -ne $null) {
            if ($workRow[$sub.WorkID] -ne $null) {
                $sub
            }
        }
    }

    $worksheet = $workbook.worksheets.Item($worksheetIndex)

    $worksheet.Name = $worksheetInfo.Name
    $worksheet.Activate()
    $workbook.Application.ActiveWindow.Zoom = $worksheetInfo.Scale

    $ColIndex = 1
    $TitleCol = $ColIndex++
    $IDCol = $ColIndex++
    $WordCountCol = $ColIndex++
    $GenreCol = $ColIndex++
    $ReadyCol = $ColIndex++
    $WithCol = $ColIndex++
    $NextCol = $ColIndex++
    $RevenueCol = $ColIndex++
    $TargetCol = $ColIndex++
    $HeaderCol = $ColIndex++
    $BaseCol = $ColIndex
    #endregion


    #
    # put in titles
    #

    Set-CellValue -Row $TitleRow         -Col $HeaderCol    -Value "Market"
    Set-CellValue -Row $IDRow            -Col $HeaderCol    -Value "ID"
    Set-CellValue -Row $URLRow           -Col $HeaderCol    -Value "Sub URL"
    Set-CellValue -Row $URLStateRow      -Col $HeaderCol    -Value "URL State"
#    Set-CellValue -Row $URLSubmissionRow -Col $HeaderCol    -Value "URL Submission"
#    Set-CellValue -Row $URLGuidelineRow  -Col $HeaderCol    -Value "URL Guideline"
#    Set-CellValue -Row $URLGeneralRow    -Col $HeaderCol    -Value "URL General"
    Set-CellValue -Row $MismatchRow      -Col $HeaderCol    -Value "Mismatch?"
    Set-CellValue -Row $ClosedRow        -Col $HeaderCol    -Value "Closed?"
    Set-CellValue -Row $OpenRow          -Col $HeaderCol    -Value "Open Date"
    Set-CellValue -Row $DeadlineRow      -Col $HeaderCol    -Value "Deadline"
    Set-CellValue -Row $RateRow          -Col $HeaderCol    -Value "Rate"
    Set-CellValue -Row $WordsRow         -Col $HeaderCol    -Value "Word Range"
    Set-CellValue -Row $TypeRow          -Col $HeaderCol    -Value "Type"
    Set-CellValue -Row $CategoryRow      -Col $HeaderCol    -Value "Category"
    Set-CellValue -Row $GenreRow         -Col $HeaderCol    -Value "Genre"
    Set-CellValue -Row $TaglineRow       -Col $HeaderCol    -Value "Tagline"
    Set-CellValue -Row $AffirmativeRow   -Col $HeaderCol    -Value "Affirmative?"
    Set-CellValue -Row $ParallelRow      -Col $HeaderCol    -Value "Parallel Subs?"
    Set-CellValue -Row $OtherRow         -Col $HeaderCol    -Value "Other guidelines"
    Set-CellValue -Row $GuideLinesRow    -Col $HeaderCol    -Value "All guidelines"
    Set-CellValue -Row $BlockedRow       -Col $HeaderCol    -Value "Blocked?"

    Set-CellValue -Row $HeaderRow        -Col $TitleCol     -Value "Work"
    Set-CellValue -Row $HeaderRow        -Col $IDCol        -Value "ID"
    Set-CellValue -Row $HeaderRow        -Col $WordCountCol -Value "Word Count"
    Set-CellValue -Row $HeaderRow        -Col $GenreCol     -Value "Genre"
    Set-CellValue -Row $HeaderRow        -Col $ReadyCol     -Value "Ready"
    Set-CellValue -Row $HeaderRow        -Col $WithCol      -Value "With"
    Set-CellValue -Row $HeaderRow        -Col $NextCol      -Value "Next"
    Set-CellValue -Row $HeaderRow        -Col $RevenueCol   -Value "Revenue"
    Set-CellValue -Row $HeaderRow        -Col $TargetCol    -Value "Target"

    #
    # set column widths
    #
    if ($GlobalFlipFlag) {
        # because we're flipped, we supply row numbers instead
        Set-ColumnProperty -Col $TitleRow  -Property "columnWidth" -Value 30
        Set-ColumnProperty -Col $IDRow     -Property "columnWidth" -Value 15
        Set-ColumnProperty -Col $URLRow    -Property "columnWidth" -Value 30
        Set-RowProperty    -Row $TargetCol -Property "rowHeight"   -Value 30
    }
    else {
        Set-ColumnProperty -Col $TitleCol  -Property "columnWidth" -Value 30
        Set-ColumnProperty -Col $GenreCol  -Property "columnWidth" -Value 20
        Set-ColumnProperty -Col $HeaderCol -Property "columnWidth" -Value 20
    }

    #
    # hide unnecessary columns
    if ($worksheetInfo.IsNovel) {
        for ($xrow = $MismatchRow; $xrow -lt $GuideLinesRow; $xrow++) {
            # Set-RowProperty -Row $xrow -Property "Hidden" -Value $true
            Set-ColumnProperty -Col $xrow -Property "Hidden" -value $true
        }
    }
    else {
        for ($xrow = $MismatchRow; $xrow -lt $DeadlineRow; $xrow++) {
            # Set-RowProperty -Row $xrow -Property "Hidden" -Value $true
            Set-ColumnProperty -Col $xrow -Property "Hidden" -value $true
        }
        for ($xrow = $TaglineRow; $xrow -lt $GuideLinesRow; $xrow++) {
            # Set-RowProperty -Row $xrow -Property "Hidden" -Value $true
            Set-ColumnProperty -Col $xrow -Property "Hidden" -value $true
        }
    }
    Set-CellMethodCall -Row 1        -Col 1        -Method "Select" -NoValue
    Set-CellMethodCall -Row $BaseRow -Col $BaseCol -Method "Select" -NoValue

    $minRow = $BaseRow - 1
    $minCol = $BaseCol - 1
    $Excel.ActiveWindow.FreezePanes = $true
    #
    # Fill-in all the market data
    #
    $maxCol = 0

    $marketSubset | foreach {
        $market = $_
        $col = $market.Index
        $xcol = $col + $BaseCol
        $maxCol = $xcol
        if (-not $GlobalFlipFlag) {
            Set-ColumnProperty                   -Col $xcol -Property "columnWidth" -Value 16
        }
        Set-CellMethodCall -Row $TitleRow    -Col $xcol -Method "Select" -NoValue
        Set-CellValue      -Row $TitleRow    -Col $xcol -Value $market.Title
        Set-CellValue      -Row $IDRow       -Col $xcol -Value $market.IDCode
        $csvInfo = $MarketCsvDataHash[$market.IDCode]
#        Set-CellProperty   -Row $URLSubmissionRow  -Col $xcol -Property "Formula" -Value "=HYPERLINK(""$($csvinfo.SubmitURL)"")"
#        Set-CellProperty   -Row $URLGuidelineRow   -Col $xcol -Property "Formula" -Value "=HYPERLINK(""$($csvinfo.GuidelineURL)"")"
#        Set-CellProperty   -Row $URLGeneralRow     -Col $xcol -Property "Formula" -Value "=HYPERLINK(""$($csvinfo.GeneralURL)"")"
        if ([string]$market.URL -eq "") {
            $market.URL = ""
            #
            # work through possibles
            if ($market.URL -eq "") {
                $market.URL = [string]$csvinfo.TestOpenURL
            }
            if ($market.URL -eq "") {
                $market.URL = [string]$csvinfo.SubmitURL
            }
            if ($market.URL -eq "") {
                $market.URL = [string]$csvinfo.GuidelineURL
            }
            if ($market.URL -eq "") {
                $market.URL = [string]$csvinfo.GeneralURL
            }
            if ($market.URL -eq "") {
                $market.URL = [string]$market.Address1
            }
            if ($market.URL -eq "") {
                $market.URL = [string]$market.Address2
            }
        }
        Set-CellProperty   -Row $URLRow      -Col $xcol -Property "Formula" -Value "=HYPERLINK(""$($market.URL)"")"
        if ($market.Guidelines -like "*Defunct*") {
            $urlState = "Defunct"
        } else {
            $MarketStatusRow = $MarketStatusLookup[$market.IDCode]
            $comments = $MarketStatusRow.SubmissionStatus -join "`r`n"
            if ($MarketStatusRow.URLList.Count -eq 0) {
                $urlState = "OK"
            }
            else {
                $urlState = "Fault"
                $badUrlList = $MarketStatusRow.URLList | foreach {
                    $URLInfo = $_
                    "$($URLInfo.URL) : $($URLInfo.ExceptionMessage)"
                }
                Set-CellMethodCall -Row $URLStateRow   -Col $xcol -Method "AddComment" -Value $($badUrlList -join "`r`n")
            }
          #  $urlState = Test-URL $market.URL -UseInvokeWR: $UseInvokeWR
            if (($urlState -eq "OK") -and (-not [string]::IsNullOrWhiteSpace($comments))) {
                Set-CellMethodCall -Row $URLStateRow   -Col $xcol -Method "AddComment" -Value $comments
            }
        }
        Set-CellValue      -Row $URLStateRow -Col $xcol -Value $urlState
        $guidelines = Analyse-Guidelines -Guidelines $market.Guidelines
        $marketState = "Open"
        $guidelines.psobject.properties | foreach {
            $n = $_.Name
            $v = $_.Value
            switch ($n) {
                "MarketType" {
                    Set-CellValue -Row $TypeRow        -Col $xcol -Value $v
                }
                "Closed" {
                    Set-CellValue -Row $ClosedRow      -Col $xcol -Value $v
                    $marketState = [string]$v
                    if ($marketState -cin @("Defunct","Ineligible")) {
                      #  Set-RowProperty -Row $xcol -Property "Hidden" -Value $true
                    }
                }
                "OpenDate" {
                    Set-CellValue -Row $OpenRow        -Col $xcol -Value $v
                }
                "Deadline" {
                    Set-CellValue -Row $DeadlineRow    -Col $xcol -Value $v
                }
                "Rate" {
                    if ([string]::IsNullOrWhiteSpace($v)) {
                        $v = $market.Rate
                    }
                    Set-CellValue -Row $RateRow        -Col $xcol -Value $v
                }
                "Genre" {
                    Set-CellValue -Row $GenreRow       -Col $xcol -Value $v
                }
                "Tagline" {
                    Set-CellValue -Row $TaglineRow     -Col $xcol -Value $v
                }
                "Parallel" {
                    Set-CellValue -Row $ParallelRow    -Col $xcol -Value $v
                }
                "Affirmative" {
                    Set-CellValue -Row $AffirmativeRow -Col $xcol -Value $v
                }
                "WordRange" {
                    Set-CellValue -Row $WordsRow       -Col $xcol -Value $v
                }
                "Category" {
                    Set-CellValue -Row $CategoryRow    -Col $xcol -Value $v
                }
                "Other" {
                    Set-CellValue -Row $OtherRow       -Col $xcol -Value $v
                }
                "" {
                    # silently ignore
                }
                default {
                    Set-CellValue -Row $OtherRow       -Col $xcol -Value $v
                    if ([string]$v -ne "") {
                        Write-Host "Unrecognised guideline keyword '$n', value '$v' (market '$($market.Title)')"
                    }
                }
            }
        }
        Set-CellMethodCall -Row $GenreRow      -Col $xcol -Method "AddComment" -Value $market.Comments
        Set-CellValue      -Row $GuideLinesRow -Col $xcol                      -Value $market.Guidelines
        if ($marketState -eq "Unreachable") {
            $ineligibleHash[$xcol] = $true
        }
        elseif ($marketState -ne $urlState) {
            if (($marketState -ne "") -or ($urlState -ne "Open")) {
                Set-CellValue    -Row $MismatchRow   -Col $xcol                      -Value "MISMATCH"
            }
        }
        elseif ($marketState -eq "Closed") {
            $ineligibleHash[$xcol] = $true
        }
    }

    #
    # Fill-in all the work data
    #
    $maxRow = 0
    $workSubset | foreach {
        $work = $_
        $row = $work.Index
        $xrow = $row + $BaseRow
        $maxRow = $xrow
        if ($GlobalFlipFlag) {
            Set-ColumnProperty            -Col $xrow         -Property "columnWidth" -Value 16
            Set-CellMethodCall -Row $xrow -Col $TitleCol     -Method   "Select"      -NoValue
        }
        Set-CellValue      -Row $xrow -Col $TitleCol                             -Value $work.Title
        Set-CellMethodCall -Row $xrow -Col $TitleCol     -Method   "AddComment"  -Value $work.Description
        Set-CellValue      -Row $xrow -Col $IDCol                                -Value $work.IDCode
        Set-CellValue      -Row $xrow -Col $WordCountCol                         -Value $work.Words
        Set-CellValue      -Row $xrow -Col $GenreCol                             -Value $work.Genre
        $comments = [string]$work.Comments
        if ($comments -match ".*Target=([\d:]+).*") {
            $targetMarketStr = [string]$Matches[1]
            $targetMarketList = $targetMarketStr -split ':'
            $targetMarketList = @($targetMarketList)
            foreach ($id in $targetMarketList) {
                $tCol = $marketIndexByIDCode[$id]
                $xtcol = $tCol + $BaseCol
                $sMarket = $marketSynthByIDCode[$id]
                Set-CellValue      -Row $xrow -Col $NextCol                                       -Value $sMarket.Title
                Set-CellMethodCall -Row $xrow -Col $xtcol         -Method "Select"                -NoValue
                Set-CellProperty   -Row $xrow -Col $xtcol         -Property "Interior.ColorIndex" -Value 22
                Set-CellProperty   -Row $xrow -Col $xtcol         -Property "Interior.Pattern"    -Value 1
            }
        }
        Set-CellValue   -Row $xrow -Col $TargetCol -Value $comments
        if ($work.Trunked -ne "0") {
            Set-CellValue -Row $xrow -Col $ReadyCol  -Value "Trunked"
        }
        else {
            Set-CellValue -Row $xrow -Col $ReadyCol  -Value "Ready"
        }
    }

    Set-CellMethodCall -Row $BaseRow -Col $BaseCol -Method "Select" -NoValue

    for ($mix = 0; $mix -lt $marketSubset.Length; $mix++) {
        #  if (-not $GlobalFlipFlag) {
        Set-CellMethodCall -Row $BaseRow -Col $($mix + $BaseCol) -Method "Select" -NoValue
        #  }
        for ($wix = 0; $wix -lt $workSubset.Length; $wix++) {
            $work = $workSubset[$wix]
            $market = $marketSubset[$mix]
            $guidelines = $marketSynthByTitle[$market.Title]
            if (($work.Category -eq $guidelines.Category) -or 
                (($work.Category -ne "Novels") -and ($guidelines.Category -ne "Novels"))) {
                # Write-Host "$($work.Title): $($work.Category) :: $($market.Title): $($guidelines.Category)"
                $CanSubmit = "Yes"
                #
                # Closed
                if ([string]$guidelines.Closed -ne "") {
                    $CanSubmit = [string]$guidelines.Closed
                }
                #
                # Unpaid
                if ([string]$guidelines.Rate -eq "Unpaid") {
                    $CanSubmit = [string]$guidelines.Rate
                }
                #
                # word limits
                if ($guidelines.WordRange -ne $null) {
                    $bounds = Split-WordRange $guidelines.WordRange
                    if ($bounds.GetType().Name -eq "Int32") {
                        Write-Host "error on split"
                    }
                    if (([int]$work.Words -lt $bounds[0]) -or ([int]$work.Words -gt $bounds[1])) {
                        $CanSubmit = "Length"
                    }
                }
                #
                # Genre - must-have
                if ($CanSubmit -eq "Yes") {
                    if (([string]$work.Genre -ne "") -and ([string]$guidelines.Genre -ne "")) {
                        $CanSubmit = "Bad Genre"
                        $work.Genre -split "\/" | foreach {
                            if ($_ -eq $guidelines.Genre) {
                                $CanSubmit = "Genre Match"
                            }
                        }
                    }
                }
                Set-CellValue -Row $($wix + $BaseRow) -Col $($mix + $BaseCol) -Value $CanSubmit
            }
        }
    }

    Set-CellMethodCall -Row $BaseRow -Col $BaseCol -Method "Select" -NoValue

    $paymentHash = @{}
    $submissionSubset | foreach {
        $submission = $_
        $su2 = $submission.Clone() | Add-Member NoteProperty Type "Submission" -PassThru
        "Comments" | foreach {
            $Field = $_
            $su2."$Field" = [System.Web.HttpUtility]::UrlDecode($su2.$Field)
        }
        $col = $marketIndexByIDCode[$su2.MarketID]
        $xcol = $col + $BaseCol
        $row = $workRow[$su2.WorkID]
        $xrow = $row + $BaseRow
        $sWork = $workHash[$su2.WorkID]
        $sMarket = $marketSynthByIDCode[$su2.MarketID]
        $agency = $AgencyInfoHash[$su2.MarketID]
        if ($agency -ne $null) {
            $agency | Out-Null
        }
        $dateBack = $su2.DateBack
        #
        # ColorIndex values - see http://dmcritchie.mvps.org/excel/colors.htm
        #                      or https://msdn.microsoft.com/en-us/library/cc296089(v=office.12).aspx#xlDiscoveringColorIndex_ColorIndexProperty
        #
        $cellData = "Sent=$($su2.DateSent)`nBack=$($dateBack)"
        if ($dateBack -eq "1899-12-29") {
            $dateBack = '????-??-??'
            $cellData = "Sent=$($su2.DateSent)`nBack=$($dateBack)"
            Set-CellValue    -Row $xrow         -Col $ReadyCol                             -Value "Out"
            Set-CellValue    -Row $xrow         -Col $WithCol                              -Value $sMarket.Title
            Set-CellValue    -Row $BlockedRow   -Col $xcol                                 -Value "BLOCKED"
            #
            # see if it is outside response parameters
            # colorindex values http://dmcritchie.mvps.org/excel/colors.htm
            $newColorIndex = 6  # yellow - it's out
            $agency | where {-not [string]::IsNullOrWhiteSpace($_.ResponsePolicy)} | Select-Object -First 1 | foreach {
                #
                # decide whether we are beyond date range
                if ($_.ResponsePolicy -match "(Response|Timeout)\s(.*)\sweeks") {
                    $type = $Matches[1]
                    $length = $Matches[2]
                    $maxtime = $length -replace "\d+\-",''
                    $sent = [datetime]::Parse($su2.DateSent)
                    $responseBy = $sent.AddDays([int]$maxtime * 7)
                    $due = $responseBy.ToString("yyyy-MM-dd")
                    #
                    # set the comment for the cell to indicate the due date
                    $feedbackText = "$type by $due)"
                    if ($type -eq "Response") {
                        $cellData = "Sent=$($su2.DateSent)`nRsp=$($due)"
                    }
                    else {
                        $cellData = "Sent=$($su2.DateSent)`nT/O=$($due)"
                    }
                    Set-CellMethodCall -Row $xrow -Col $xcol     -Method   "AddComment"  -Value $feedbackText
                    if ($today -gt $responseBy) {
                        $newColorIndex = 46   # dark orange - out of range
                    }
                }
            }
            Set-CellProperty -Row $xrow         -Col $xcol -Property "Interior.ColorIndex" -Value $newColorIndex
            Set-CellProperty -Row $xrow         -Col $xcol -Property "Interior.Pattern"    -Value 1
            $ineligibleHash[$xcol] = $true
            $outHash[$xrow] = $true
        }
        if ([string]$su2.Published -ne "0") {
            Set-CellValue    -Row $xrow         -Col $ReadyCol                             -Value "Published"
            Set-CellValue    -Row $xrow         -Col $xcol                                 -Value "Published"
            Set-CellProperty -Row $xrow         -Col $xcol -Property "Interior.ColorIndex" -Value 4
            Set-CellProperty -Row $xrow         -Col $xcol -Property "Interior.Pattern"    -Value 1
        }
        elseif ([string]$su2.Sale -ne "0") {
            Set-CellValue    -Row $xrow         -Col $ReadyCol                             -Value "Sold"
            Set-CellValue    -Row $xrow         -Col $xcol                                 -Value "Sold"
            Set-CellProperty -Row $xrow         -Col $xcol -Property "Interior.ColorIndex" -Value 50
            Set-CellProperty -Row $xrow         -Col $xcol -Property "Interior.Pattern"    -Value 1
        }
        else {
            Set-CellValue -Row $xrow -Col $xcol -Value $cellData
        }
        $payment = [double]$su2.Payment
        $paymentHash[$su2.WorkID] += $payment
        if ($payment -gt 0) {
            Set-CellValue    -Row $xrow         -Col $RevenueCol -Value $paymentHash[$su2.WorkID]
        }

        $market = $marketSynthByIDCode[$su2.MarketID]

        $work = $workHash[$su2.WorkID]

        $key = "$($su2.WorkID)|$($su2.MarketID)"
        $submissionHash[$key] = $su2
    }

    #
    # turn on autofilter
    #
    if ($false) {
        $range = $worksheet.UsedRange
        #  $range.EntireColumn.AutoFilter() | Out-Null
    }
    else {
        if ($GlobalFlipFlag) {
            $firstRow = $minCol
            $lastRow = $maxCol
            $firstCol = 1
            $lastCol = $maxRow
        }
        else {
            $firstRow = $minRow
            $lastRow = $maxRow
            $firstCol = 1
            $lastCol = $maxCol
        }
        $str = "R${firstRow}C${firstCol}:R${lastRow}C${lastCol}"
        $str2 = $app.ConvertFormula($str, $xlR1C1, $xlA1)
        $range = $worksheet.Range($str2)
        $range.Select() | Out-Null
    }
    $range.AutoFilter() | Out-Null

    #
    # hide ineligible columns
    if ($HideIneligible -and -not $GlobalFlipFlag) {
        foreach ($item in $ineligibleHash.GetEnumerator()) {
            Set-ColumnProperty -Col $item.Key -Property "Hidden" -Value $true
        }
    }
    #
    # hide "out" rows
    if ($HideOut -and $GlobalFlipFlag) {
        foreach ($item in $outHash.GetEnumerator()) {
            Set-ColumnProperty -Col $item.Key -Property "Hidden" -Value $true
        }
    }
    #
    # leave the upper left cell selected
    Set-CellMethodCall -Row $BaseRow -Col $BaseCol -Method "Select" -NoValue

}

#
# All rows copied - close the spreadsheet
#
$workbook.saveas($SubmissionFile)
$Excel.Quit()
Remove-Variable -Name Excel
[gc]::collect()
[gc]::WaitForPendingFinalizers()

Out-File ".\${JobName}_Excel_Done.timestamp"

Write-Host "URL Issues"
Write-Host "=========="

$connectionErrors.GetEnumerator() | foreach {
    New-Object PSObject |
        Add-Member NoteProperty Status $_.Value -PassThru |
        Add-Member NoteProperty URL    $_.Key   -PassThru
} | Sort-Object -Property Status,URL | Format-Table

Write-Host "Missing Markets"
Write-Host "==============="

$noMarketData.GetEnumerator() | foreach {
    New-Object PSObject |
        Add-Member NoteProperty SonarID    $_.Key   -PassThru |
        Add-Member NoteProperty Title      $_.Value -PassThru
} | Sort-Object -Property Title,SonarID | Format-Table

Write-Host "Done"
