#
# script to analyse the latest sonar file & report on current status
# See: https://adamtheautomator.com/powershell-excel-tutorial/
#      https://kpatnayakuni.com/2019/01/13/excel-reports-using-importexcel-module-from-powershell-gallery/
#

param ($JobName = "Local",
       $SubmissionData = 'P:\Users\Bill\Documents\My Stories\Submissions\All Submissions.sonar3')

Add-Type -AssemblyName system.web

Import-Module ImportExcel

Import-Module PowerHTML

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
$SubmissionScheduleFile       = "$sdir\SubmissionSchedule.json"
#endregion

#
# see https://www.cogsci.ed.ac.uk/~richard/utf-8.cgi?input=EF+BF+BD&mode=bytes
function Convert-ANSIEncodedString ($InputString) {
    $target = ""
    while ($InputString -match "^([^%]*)(%[A-F0-9]{2})(.*)$") {
       $target += $Matches[1]
       $chy = $Matches[2] -replace '%','0x'
       $chx = [Text.Encoding]::Default.GetString([byte[]] $chy)
       $target += $chx
       $InputString = $Matches[3]
    }
    $target += $InputString
    $target
}

$test = "Th%E9r%E8se%20Coen"

$result = Convert-ANSIEncodedString -InputString $test

#
# convert accented characters to non-accented characters
# see https://cosmoskey.blogspot.com/2009/09/powershell-function-convert.html
function Convert-DiacriticCharacters {
    param(
        [string]$inputString
    )
    [string]$formD = $inputString.Normalize(
            [System.text.NormalizationForm]::FormD
    )
    if ($inputString -like "*Coen*") {
        $inputString | Out-Null
    }
    $stringBuilder = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $formD.Length; $i++){
        $unicodeCategory = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($formD[$i])
        $nonSPacingMark = [System.Globalization.UnicodeCategory]::NonSpacingMark
        if($unicodeCategory -ne $nonSPacingMark){
            $stringBuilder.Append($formD[$i]) | Out-Null
        }
    }
    $stringBuilder.ToString().Normalize([System.text.NormalizationForm]::FormC)
}

$script:OnSubmission = 0
$script:Acceptances = 0
$script:Rejections = 0

$AgentsByMarketHash = @{}
$AllAgents = Import-Excel -Path $individualAgentsSheetFile
$AllAgents | foreach {
    $AgentRow = $_
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
}

$SonarIDLookup = @{}

if (Test-Path -LiteralPath $RalanMarketsFile) {
    Import-Csv -LiteralPath $RalanMarketsFile | where {-not [string]::IsNullOrWhiteSpace($_.SonarID)} | foreach {
        $RalanRec = $_
        $SonarIDLookup[$RalanRec.UniqueName] = $RalanRec
        $SonarIDLookup[$RalanRec.SonarID] = $RalanRec
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

$MarketFields = "Name,SonarID,HomeURL,Subs,WorkBounds,Agent,Genre,Out,Back,SLA,SLADate" -split ','

function New-Market () {
    New-Object PSObject | Select-Object -Property $MarketFields
}

#region Build from Sonar File
if ($false) {
    $xmlReader = New-Object -TypeName XML
    $doc = $xmlReader.Load($SubmissionData)
}
else {
    $SubXml = Get-Content -LiteralPath $SubmissionData

    $doc = [xml]$SubXml
}

$root = $doc.SONAR3

$MarketsHash = @{}
$root.MARKETS.MARKET | foreach {
    $marketXML = $_
    $MarketsHash[$marketXML.IDCode] = $marketXML
}

function Get-TitleAbbreviation ($Title) {
    $TheWords = [object[]]($Title -replace "[^A-Z\s]",'' -split "\s+")
    $Abbrev = ""
    $TheWords | % {$Abbrev += $_.SubString(0,1)}
    if ($Abbrev.Length -lt 3) {
        $Pad = $TheWords[$TheWords.Count - 1] + "AAA"
        $Abbrev += $Pad.SubString(1,3 - $Abbrev.Length)
    }
    $Abbrev.ToUpper()
}

$WorksHash = @{}
$root.WORKS.WORK | foreach {
    $workXML = $_
    $fullTitle = Convert-ANSIEncodedString $workXML.Title
    $PrincipalTitle = $fullTitle -replace "\/.*",'' -replace "\s+$",''
    $WorkPrefix = Get-TitleAbbreviation -Title $PrincipalTitle
    $el = $doc.CreateElement("Prefix")
    $el.set_InnerText($WorkPrefix) | Out-Null
    $workXML.AppendChild($el) | Out-Null
    $WorksHash[$workXML.IDCode] = $workXML
}

$BlockedMarketsHash = @{}
$SubmissionsHash = @{}
$root.SUBMISSIONS.SUBMISSION | foreach {
    $submissionXML = $_
    if ($submissionXML.DateBack -eq "1899-12-29") {
        $BlockedMarketsHash[$submissionXML.MarketID] = $true
    }
    elseif ($submissionXML.DateBack -notmatch "\d\d\d\d\-\d\d\-\d\d") {
        $BlockedMarketsHash[$submissionXML.MarketID] = $true
    }
    $SubmissionsHash[$submissionXML.IDCode] = $submissionXML
}

$worksByLength = $WorksHash.GetEnumerator() | % {$_.Value} | foreach {
    $workXml = $_
    $Title = [System.Web.HttpUtility]::UrlDecode($workXml.Title)
    if ($Title -like "*Beans*") {
        $workXml | Out-Null
    }
    $BestAgents = @()
    $BestAgencies = @()
    [System.Web.HttpUtility]::UrlDecode($workXML.Comments) -split "`r`n" | foreach {
        $line = $_
        if ($line -match "^(.*) at (.*)$") {
            $agentName = $Matches[1]
            $agency = $Matches[2]
            $BestAgents += (Convert-ANSIEncodedString -InputString $agentName)
            $BestAgencies += $agency
        }
    }
    $Genre = [System.Web.HttpUtility]::UrlDecode($workXml.Genre)
    $WorkLength = [int]($workXml.Words)
    $IsNovel = $WorkLength -gt 40000
    $IsTrunked = ($workXml.Trunked -ne '0')
    New-Object PSObject |
        Add-Member NoteProperty WorkName             $Title              -PassThru |
        Add-Member NoteProperty WorkID               $workXml.IDCode     -PassThru |
        Add-Member NoteProperty Genre                $Genre              -PassThru |
        Add-Member NoteProperty IsTrunked            $IsTrunked          -PassThru |
        Add-Member NoteProperty IsNovel              $IsNovel            -PassThru |
        Add-Member NoteProperty WorkLength           $WorkLength         -PassThru |
        Add-Member NoteProperty WorkGroup            0                   -PassThru |
        Add-Member NoteProperty SubmissionMarketIDs  @{}                 -PassThru |
        Add-Member NoteProperty SalesMarketIDs       @{}                 -PassThru |
        Add-Member NoteProperty FirstSubmissionDate  $null               -PassThru |
        Add-Member NoteProperty SubmissionHash       @{}                 -PassThru |
        Add-Member NoteProperty ActiveSubmissions    0                   -PassThru |
        Add-Member NoteProperty Sales                0                   -PassThru |
        Add-Member NoteProperty BestAgents           $BestAgents         -PassThru |
        Add-Member NoteProperty BestAgencies         $BestAgencies       -PassThru |
        Add-Member NoteProperty Prefix               $workXml.Prefix     -PassThru
} | Sort-Object -Property WorkLength,WorkName

$allMarkets = $MarketsHash.GetEnumerator() | % {$_.Value} | foreach {
    $marketXML = $_
    if ($marketXML.IDCode -in @("67795","5350","26137")) {
        $market | Out-Null
    }
    $RalanRec = $SonarIDLookup[$marketXML.IDCode]
    $MarketName = [System.Web.HttpUtility]::UrlDecode($marketXML.Title)
    $KeyWords = [System.Web.HttpUtility]::UrlDecode($marketXML.Guidelines) -split ','
    $IsNovel = "Novels" -in $KeyWords
    $IsDefunct = "Defunct" -in $KeyWords
    $IsPro = "SFWA" -in $KeyWords
    $validURLs = @()
    if ($RalanRec -ne $null) {
        if ($RalanRec.MarketType -eq "Pro") {
            $IsPro = $true
        }
    }
    if ($IsPro) {
        $marketXML | Out-Null
        "Email,URL,Address1,Address2,Address3,Address4" -split ',' | foreach {
            $fieldName = $_
            $marketXML.$fieldName
        } | where {-not [string]::IsNullOrWhiteSpace($_)} | where {$_ -like "http*"} | 
        foreach {
            $url = [System.Web.HttpUtility]::UrlDecode($_)
            $url
        } | Sort-Object -Unique | foreach {
            $validURLs += $_
        }
    }
    $WordBounds = $null
    $Genres = @{}
    switch -Regex ($KeyWords) {
        "^(\d+)\-(\d+)$" {
                $min = [int]($Matches[1])
                $max = [int]($Matches[2])
                $WordBounds = New-Object PSObject |
                    Add-Member NoteProperty MinVal $min -PassThru |
                    Add-Member NoteProperty MaxVal $max -PassThru
            }
        "^(Fantasy|Horror|SF|SFF|SpecFic|Steampunk|Noir|Upbeat|Dark|Humour|Crime|Mystery|Historical|Romance|Literary|Contemporary|Christian|Thriller)$" {
                $Genres[$Matches[1]] = $true
            }
    }
    New-Object PSObject |
        Add-Member NoteProperty MarketName $MarketName         -PassThru |
        Add-Member NoteProperty MarketID   $marketXML.IDCode   -PassThru |
        Add-Member NoteProperty IsNovel    $IsNovel            -PassThru |
        Add-Member NoteProperty IsDefunct  $IsDefunct          -PassThru |
        Add-Member NoteProperty IsPro      $IsPro              -PassThru |
        Add-Member NoteProperty WordBounds $WordBounds         -PassThru |
        Add-Member NoteProperty KeyWords   $KeyWords           -PassThru |
        Add-Member NoteProperty Genres     $Genres             -PassThru |
        Add-Member NoteProperty Rules      @()                 -PassThru |
        Add-Member NoteProperty AllURLs    $validURLs          -PassThru
}

$activeMarkets = $allMarkets | where {-not $_.IsDefunct}

$proMarkets = $activeMarkets | where {$_.IsPro}

#$proMarkets | Select-Object MarketName,MarketID,AllURLs,SubmissionURL,Rules | Sort-Object -Property MarketName | ConvertTo-Json | Set-Content -Encoding UTF8 -LiteralPath $SubmissionScheduleFile

#endregion

if (Test-Path -LiteralPath $NewSubmissionFile) {
    Remove-Item -LiteralPath $NewSubmissionFile
}

#
# prioritise the order on which we see the stories
$worksByLength | Sort-Object -Property IsTrunked,WorkLength,WorkName | foreach {
    $work = $_
    $WorkSubmissionHash = @{}
    $Sales = @{}
    $submissionList = $SubmissionsHash.GetEnumerator() | % {$_.Value} | where { $_.WorkID -eq $work.WorkID } | foreach {
        $submissionXML = $_
        if ($submissionXML.Sale -notmatch '0') {
            $script:Acceptances += 1
        }
        elseif ($submissionXML.DateBack -eq "1899-12-29") {
            $script:OnSubmission += 1
        }
        elseif ($submissionXML.DateBack -match "\d\d\d\d\-\d\d\-\d\d") {
            $script:Rejections += 1
        }
        else {
            $script:OnSubmission += 1
        }
        $work.SubmissionMarketIDs[$submissionXML.MarketID] = $true
        if ($work.FirstSubmissionDate -eq $null) {
            $work.FirstSubmissionDate = $submissionXML.DateSent
        }
        else {
            if ($submissionXML.DateSent -lt $work.FirstSubmissionDate) {
                $work.FirstSubmissionDate = $submissionXML.DateSent
            }
        }
        $WorkSubmissionHash[$submissionXML.IDCode] = $submissionXML
        if ($submissionXML.Sale -ne '0') {
            $Sales[$submissionXML.IDCode] = $submissionXML
            $work.SalesMarketIDs[$submissionXML.MarketID] = $true
        }
        $submissionXML
    }
    if ($submissionList -ne $null) {
        $submissionList = @($submissionList)
        $submissionList | foreach {
            $submissionXML = $_
            $work.SubmissionHash[$submissionXML.MarketID] = $submissionXML
        }
        if ($work.SalesMarketIDs.Count -gt 0) {
            Write-Host @debugColours "$($work.WorkName) has been submitted $($work.SubmissionMarketIDs.Count) times and sold $($work.SalesMarketIDs.Count) times"
            $work.WorkGroup = 4
        }
        else {
            Write-Host @warningColours "$($work.WorkName) has been submitted $($work.SubmissionMarketIDs.Count) times"
            if ($work.IsNovel) {
                $work.WorkGroup = 2
            }
            else {
                $work.WorkGroup = 3
            }
        }
    }
    else {
        Write-Host @errorColours "$($work.WorkName) has never been submitted"
        $work.WorkGroup = 1
    }
    if ($work.IsTrunked) {
        $work.WorkGroup = 5
    }
    #
    #
}

$PriorityList = @{}

"***SOLD***,Yes,Open,,Closed,BLOCKED,Bad Length,On Submission,Expected,Rejected" -split ',' | foreach {
    $str = $_
    $Priority = $PriorityList.Count
    $PriorityList[$str] = $Priority
}

function Get-SubsPrecedence ($SubsString) {
    if ([string]::IsNullOrWhiteSpace($SubsString)) {
        $SubsString = ""
    }
    if ($SubsString.StartsWith("Closed")) {
        $SubsString = "Closed"
    }
    if ($SubsString.StartsWith("Open")) {
        $SubsString = "Open"
    }
    if ($SubsString.StartsWith("BLOCKED")) {
        $SubsString = "BLOCKED"
    }
    $priority = $PriorityList[$SubsString]
    if ($priority -eq $null) {
        $priority = 99
        Write-Host @warningColours "Submission Type $SubsString not sortable"
    }
    $priority
}

$SpareRows = 5

$MarketSchedules = @{}

$sj3 = Get-Content -LiteralPath $SubmissionScheduleFile | ConvertFrom-Json
$sj3 | foreach {
   $marketScheduleRow = $_
   $MarketSchedules[$marketScheduleRow.MarketID] = $marketScheduleRow
}

$today = Get-Date

#region Market Window Detail
function New-DateBand {
    New-Object PSObject | Select-Object From,To,Status
}

$DateTableHash = @{}

function Get-DateTable ($URL) {
    $dateTable = $DateTableHash[$URL]
    if ($dateTable -eq $null) {
        $info = Invoke-WebRequest -Uri $URL -TimeoutSec 20 -UseBasicParsing | ConvertFrom-Html
        $textToParse = $info.InnerHtml -split "[`r`n]" |
            where {-not [string]::IsNullOrWhiteSpace($_)} | 
            where {$_ -match "^\s*<[\/]{0,1}t.*$"}
        $fieldList = @()
        $stack = New-Object System.Collections.Queue
        $index = 0
        $item = New-DateBand
        $stack.Enqueue($item)
        $textToParse | foreach {
            $line = $_ -replace "^\s*",'' -replace "\s*$",''
            if ($line -match "<th>(.*)<\/th>") {
                $fieldList += $Matches[1]
            }
            elseif ($line -match "<td>(.*)<\/td>") {
                $fieldVal = $Matches[1]
                $fieldName = $fieldList[$index]
                $item.$fieldName = $fieldVal
                $index++
                if ($index -eq $fieldList.Count) {
                    $index = 0
                    $item = New-DateBand
                    $stack.Enqueue($item)
                }
            }
            else {
            }
        }
        $dateTable = $stack.GetEnumerator() | where {-not [string]::IsNullOrWhiteSpace($_.Status)} | foreach {
            $item = $_
            $item.From = [datetime]::Parse($item.From)
            $item.To = [datetime]::Parse($item.To)
            $item
        }
        $DateTableHash[$URL] = $dateTable
    }
    $dateTable
}

function Get-MarketOpenState ($MarketRec) {
    $MarketStatusRec = $MarketStatusLookup[$market.MarketID]
    $marketScheduleRow = $MarketSchedules[$market.MarketID]
    if ($MarketStatusRec -eq $null) {
        return $null
    }
    if ($marketScheduleRow -ne $null) {
        $marketScheduleRow | Out-Null
        if ($marketScheduleRow.Rules.Count -gt 0) {
            foreach ($rule in $marketScheduleRow.Rules) {
                switch -regex ($rule.RuleType) {
                    "IndefiniteStatus" {
                            return "$($rule.Status) sine die"
                        }
                    "PermanentStatus" {
                            return "$($rule.Status)"
                        }
                    "OpenMonths|PartMonths" {
                            if ($today.ToString("MMMM") -in $rule.MonthList) {
                                if ($rule.RuleType -eq "OpenMonths") {
                                    return "Open"
                                }
                                else {
                                    $firstLast = $rule.DayRange -split '\-' | foreach {[int]$_}
                                    if (($today.Day -ge $firstLast[0]) -and ($today.Day -le $firstLast[1])) {
                                        return "Open"
                                    }
                                    else {
                                        return "Closed"
                                    }
                                }
                            }
                            else {
                                return "Closed - open in $($rule.MonthList -join ',')"
                            }
                        }
                    "URLTable" {
                            $dateTable = Get-DateTable -URL $rule.URL
                            $dateTable | foreach {
                                $dateBand = $_
                                if (($dateBand.From -le $today) -and ($dateBand.To.AddDays(1) -gt $today)) {
                                    return "$($dateBand.Status) until $($dateBand.To.AddDays(1).ToString("yyyy-MM-dd"))"
                                }
                            }
                        }
                    "SingleWindow" {
                            $OpenDate = [datetime]::Parse($rule.StartDate)
                            $ClosedDate = [datetime]::Parse($rule.EndDate)
                            if ($today -lt $OpenDate) {
                                return "Closed until $($rule.StartDate)"
                            }
                            elseif ($today -gt $ClosedDate.AddDays(1)) {
                                return "Closed since $($ClosedDate.AddDays(1).ToString("yyyy-MM-dd"))"
                            }
                            else {
                                return "Open till $($ClosedDate.AddDays(1).ToString("yyyy-MM-dd"))"
                            }
                        }
                    "OpenFrom|NextOpen" {
                            $OpenDate = [datetime]::Parse($rule.Date)
                            if ($today -gt $OpenDate) {
                                return "Open"
                            }
                            else {
                                return "Closed until $($rule.Date)"
                            }
                        }
                    default {
                            Write-Host "$($rule.RuleType)"
                        }
                }
            }
        }
    }
#    return $null
}

#endregion

function Export-Markets ([array]$MarketList,
                         [string]$MarketName,
                         [string]$WorksheetName,
                         [int]$StartRow,
                         $Work) {
    if ($MarketList.Count -eq 0) {
        return 0
    }
    else {
        if ($WorksheetName.Length -gt 28) {
            $WorksheetName = $WorksheetName.Substring(0,28)
        }
        Write-Host "$WorksheetName : $MarketName"
        $marketSubset = $MarketList | foreach {
            $market = $_
            if ($market.MarketID -in @("172470804" <# ,"67795","5350" #>)) {
                $market | Out-Null
            }
            $marketOut = New-Market
            if ("Childrens" -in $market.KeyWords) {
                $marketOut.Genre = "Childrens"
            }
            $marketXML = $MarketsHash[$market.MarketID]
            $AgentCollection = $AgentsByMarketHash[$market.MarketID]
            switch -wildcard ($Work.WorkName) {
                "Black Tye*" {
                        $FieldName = "Tin2Gold"
                    }
                "Thaumatist*" {
                        $FieldName = "Thaumatist"
                    }
                "John Does Blues*" {
                        $FieldName = "Dom"
                    }
                "Strike*" {
                        $FieldName = "BobbysDawn"
                    }
                default {
                        $FieldName = "N/A"
                    }
                
            }
            $InterestedAgents = @()
            #
            # work out if there are any preferred agents
            $preferredAgents = @()
            $otherAgents = @()
            #
            $AgentCollection | where {$_ -ne $null} | foreach {
                $AgentRow = $_
                $fieldVal = $AgentRow.$FieldName
                switch -Regex ($fieldVal) {
                    "^(Y|Send)$" {
                            if ($AgentRow.BestContact -in $Work.BestAgents) {
                                $preferredAgents += (Convert-ANSIEncodedString -InputString $AgentRow.BestContact)
                            }
                            else {
                                $otherAgents += (Convert-ANSIEncodedString -InputString $AgentRow.BestContact)
                            }
                        }
                }
            }
            $marketOut.Agent = Convert-ANSIEncodedString -InputString "$($preferredAgents -join ',')($($otherAgents -join ','))"
            $marketOut.Name = $market.MarketName
            $marketOut.SonarID = $market.MarketID
            $marketOut.HomeURL = [System.Web.HttpUtility]::UrlDecode($marketXML.URL)
            if ($market.MarketID -in @("48989" <# ,"67795","5350" #>)) {
                $market | Out-Null
            }
            $definitiveMarketStatus = Get-MarketOpenState -MarketRec $market | where {$_ -ne $null} | Select-Object -Last 1
            $MarketStatusLookup[$market.MarketID] | foreach {
                $MarketStatus = $_
                if ($definitiveMarketStatus -ne $null) {
                    $marketOut.Subs = $definitiveMarketStatus
                }
                else {
                    if ([string]::IsNullOrWhiteSpace($marketOut.Subs)) {
                        if ($MarketStatus.SentimentHash.Closed -eq $true) {
                            $marketOut.Subs = "Closed"
                        }
                        elseif ($MarketStatus.SentimentHash.Open -eq $true) {
                            $marketOut.Subs = "Open"
                        }
                        elseif ($MarketStatus.SentimentHash.Future -eq $true) {
                            $marketOut.Subs = "Expected"
                        }
                    }
                }
                if ($market.WordBounds -ne $null) {
                    $marketOut.WorkBounds = "$($market.WordBounds.MinVal)-$($market.WordBounds.MaxVal)"
                    if (($Work.WorkLength -lt $market.WordBounds.MinVal) -or ($Work.WorkLength -gt $market.WordBounds.MaxVal)) {
                        $marketOut.Subs = "Bad Length"
                    }
                }
            }
            if ($BlockedMarketsHash[$market.MarketID] -eq $true) {
                if ($MarketName -eq "Submitted") {
                    $marketOut.Subs = "On Submission"
                    $Work.ActiveSubmissions++
                }
                else {
                    $marketOut.Subs = "BLOCKED"
                }
            }
            if ($market.WordBounds -ne $null) {
                $marketOut.WorkBounds = "$($market.WordBounds.MinVal)-$($market.WordBounds.MaxVal)"
                if ([string]::IsNullOrWhiteSpace($marketOut.Subs) -or ($marketOut.Subs -in @("Open","Yes","Expected"))) {
                    if (($Work.WorkLength -ge $market.WordBounds.MinVal) -and ($Work.WorkLength -le $market.WordBounds.MaxVal)) {
                        $marketOut.Subs = "Yes"
                    }
                    else {
                        $marketOut.Subs = "Bad Length"
                    }
                }
            }
            $work.SubmissionHash[$market.MarketID] | where {$_ -ne $null} | foreach {
                $submissionXML = $_
                $AgentName = (Convert-ANSIEncodedString -InputString $submissionXML.Comments) -split "[`r`n]+" | where {$_ -match "^Agent\s*=\s*.*$"} | foreach {
                    if ($_ -match "^Agent\s*=\s*(.*)$") {
                        $Matches[1] -replace "^\s*",'' -replace "\s*$",'' -replace "\s+",' '
                    }
                }
                if ($AgentName -ne $null) {
                    $marketOut.Agent = Convert-ANSIEncodedString -InputString $AgentName
                }
                else {
                    $AgentCollection | where {$_ -ne $null} | foreach {
                         $AgentRow = $_
                         $action = $AgentRow.$FieldName
                         switch ($action) {
                             "Sent" {
                                     $marketOut.Agent = Convert-ANSIEncodedString -InputString $AgentRow.BestContact
                                 }
                             "Declined" {
                                     $marketOut.Agent = Convert-ANSIEncodedString -InputString $AgentRow.BestContact
                                 }
                             default {
                                     $AgentRow.BestContact | Out-Null
                                 }
                         }
                    }
                }
                $marketOut.Out = $submissionXML.DateSent
                #
                # let's work out the SLA
                $marketOut.SLA = $AgentRow.ResponsePolicy
                if ([string]::IsNullOrWhiteSpace($marketOut.SLA)) {
                    #
                    # let's find it in the market
                    if (-not [string]::IsNullOrWhiteSpace($market.MarketID)) {
                        $MarketStatusRow = $MarketStatusLookup[$market.MarketID]
                        if ($MarketStatusRow -ne $null) {
                            $marketOut.SLA = $MarketStatusRow.ResponsePolicy
                        }
                    }
                }
                if (-not [string]::IsNullOrWhiteSpace($marketOut.SLA)) {
                    $SLAItems = $marketOut.SLA -split "[\-\s]+" | Select-Object -Last 2
                    switch ($SLAItems[0]) {
                        "a" {
                                $periodQuantity = [int]1
                            }
                        "one" {
                                $periodQuantity = [int]1
                            }
                        "two" {
                                $periodQuantity = [int]2
                            }
                        "three" {
                                $periodQuantity = [int]3
                            }
                        "four" {
                                $periodQuantity = [int]4
                            }
                        "five" {
                                $periodQuantity = [int]5
                            }
                        "six" {
                                $periodQuantity = [int]6
                            }
                        "seven" {
                                $periodQuantity = [int]7
                            }
                        "eight" {
                                $periodQuantity = [int]8
                            }
                        "nine" {
                                $periodQuantity = [int]9
                            }
                        "ten" {
                                $periodQuantity = [int]10
                            }
                        "eleven" {
                                $periodQuantity = [int]11
                            }
                        "twelve" {
                                $periodQuantity = [int]12
                            }
                        "thirty" {
                                $periodQuantity = [int]30
                            }
                        "sixty" {
                                $periodQuantity = [int]60
                            }
                        default {
                                $periodQuantity = [int]$_
                            }
                    }
                    switch -Wildcard ($SLAItems[1]) {
                        "day*" {
                                $marketOut.SLADate = [datetime]::Parse($marketOut.Out).AddDays($periodQuantity).ToString("yyyy-MM-dd")
                            }
                        "week*" {
                                $marketOut.SLADate = [datetime]::Parse($marketOut.Out).AddDays(7 * $periodQuantity).ToString("yyyy-MM-dd")
                            }
                        "month*" {
                                $marketOut.SLADate = [datetime]::Parse($marketOut.Out).AddMonths($periodQuantity).ToString("yyyy-MM-dd")
                            }
                    }

                }
                $marketOut.Subs = "On Submission"
                if ($submissionXML.DateBack -ne "1899-12-29") {
                    $marketOut.Back = $submissionXML.DateBack
                    $marketOut.Subs = "Rejected"
                }
                if ($submissionXML.Sale -gt '0') {
                    $marketOut.Subs = "***SOLD***"
                }
            }
            if ($market.Genres.Count -gt 0) {
                $marketOut.Genre = ($market.Genres.GetEnumerator() | % {$_.Key}) -join ','
            }
            if ($market.IsDefunct) {
                $marketOut.Subs = "Defunct"
            }
            $marketOut
        }
        do {
            try {
                if ($MarketName -ne "Submitted") {
                    $TableName = "ZZ_" + $Work.Prefix + $MarketName
                }
                else {
                    $TableName = $WorksheetName
                }
                $TableName = $($TableName -replace "\s+",'' -replace "\p{P}+",'')
                $marketSubset | Sort-Object -Property @{Expression="Out"; Descending=$true},@{Expression = {Get-SubsPrecedence $_.Subs}; Descending=$false},@{Expression="Name"; Descending=$false} |
                                Export-Excel -Path $NewSubmissionFile `
                                             -AutoSize `
                                             -AutoFilter `
                                             -WorksheetName $WorksheetName `
                                             -StartRow $StartRow `
                                             -StartColumn 1 `
                                             -TableName $TableName `
                                             -Title "Markets: $MarketName"
                $failed = $false
            }
            catch {
                $failed = $true
            }
        } while ($failed)
        return $MarketList.Count + $SpareRows
    }
}

$FixSet = @{}

#
# now generate the spreadsheet report, with the stories I care most about at the front
$worksInOrder = $worksByLength | Sort-Object -Property @{Expression="WorkGroup";Descending=$false},@{Expression="FirstSubmissionDate";Descending=$true}
$worksInOrder | foreach {
    $work = $_
    $workShortName = $work.WorkName -replace "\s*\/.*$",''
    $marketsForWork = $activeMarkets | where {$_.IsNovel -eq $work.IsNovel}
    $startRow = 12
    if ($work.IsNovel) {
        $Sold = @()
        $Submitted = @()
        $BestAgents = @()
        $Agents = @()
        $Publishers = @()
        $marketsForWork | foreach {
            $market = $_
            if ($work.SalesMarketIDs[$market.MarketID] -eq $true) {
                $Sold += $market
                $work.Sales++
            }
            elseif ($work.SubmissionMarketIDs[$market.MarketID] -eq $true) {
                $Submitted += $market
            }
            elseif ("Agency" -in $market.KeyWords) {
               # Write-Host $market.MarketName
                if ($market.MarketName -in $work.BestAgencies) {
                    $BestAgents += $market
                }
                else {
                    $Agents += $market
                }
            }
            else {
                $Publishers += $market
            }
            #$marketXML = $MarketsHash[$market.MarketID]
        }
        #$startRow = 2

        $rowsUsed = Export-Markets -MarketList $Sold `
                                   -MarketName "Sold" `
                                   -WorksheetName $workShortName `
                                   -StartRow $startRow `
                                   -Work $work
        $startRow += $rowsUsed

        $rowsUsed = Export-Markets -MarketList $BestAgents `
                                   -MarketName "BestAgents" `
                                   -WorksheetName $workShortName `
                                   -StartRow $startRow `
                                   -Work $work
        $startRow += $rowsUsed

        $rowsUsed = Export-Markets -MarketList $Agents `
                                   -MarketName "Agents" `
                                   -WorksheetName $workShortName `
                                   -StartRow $startRow `
                                   -Work $work
        $startRow += $rowsUsed

        $rowsUsed = Export-Markets -MarketList $Publishers `
                                   -MarketName "Publishers" `
                                   -WorksheetName $workShortName `
                                   -StartRow $startRow `
                                   -Work $work
        $startRow += $rowsUsed

        $rowsUsed = Export-Markets -MarketList $Submitted `
                                   -MarketName "Submitted" `
                                   -WorksheetName $workShortName `
                                   -StartRow $startRow `
                                   -Work $work
        #
        if ($rowsUsed -gt 0) {
            # need to put the formula '="out "&DATEDIF(H105,TODAY(),"d")&" days"' in all cells where the out column is a date and the back column is blank
            $fixObj = New-Object PSObject |
                Add-Member NoteProperty WorksheetName $workShortName           -PassThru |
                Add-Member NoteProperty StartRow      $startRow                -PassThru |
                Add-Member NoteProperty RowCount      ($rowsUsed - $SpareRows) -PassThru
            $FixSet[$workShortName] = $fixObj
        }
        $startRow += $rowsUsed

    }
    else {
        $Sold = @()
        $Pro = @()
        $Other = @()
        $Submitted = @()
        $marketsForWork | foreach {
            $market = $_
            $RalanRec = $SonarIDLookup[$market.MarketID]
            if ($work.SalesMarketIDs[$market.MarketID] -eq $true) {
                $Sold += $market
            }
            elseif ($work.SubmissionMarketIDs[$market.MarketID] -eq $true) {
                $Submitted += $market
            }
            elseif ("SFWA" -in $market.KeyWords) {
                $Pro += $market
            }
            elseif ($RalanRec -eq $null) {
                $Other += $market
            }
            elseif ($RalanRec.MarketType -eq "Pro") {
                $Pro += $market
            }
            else {
                $Other += $market
            }
        }
        #$startRow = 2

        $rowsUsed = Export-Markets -MarketList $Sold `
                                   -MarketName "Sold" `
                                   -WorksheetName $workShortName `
                                   -StartRow $startRow `
                                   -Work $work
        $startRow += $rowsUsed

        $rowsUsed = Export-Markets -MarketList $Pro `
                                   -MarketName "Pro" `
                                   -WorksheetName $workShortName `
                                   -StartRow $startRow `
                                   -Work $work
        $startRow += $rowsUsed

        $rowsUsed = Export-Markets -MarketList $Other `
                                   -MarketName "Other" `
                                   -WorksheetName $workShortName `
                                   -StartRow $startRow `
                                   -Work $work
        $startRow += $rowsUsed

        $rowsUsed = Export-Markets -MarketList $Submitted `
                                   -MarketName "Submitted" `
                                   -WorksheetName $workShortName `
                                   -StartRow $startRow `
                                   -Work $work
        #
        if ($rowsUsed -gt 0) {
            # need to put the formula '="out "&DATEDIF(H105,TODAY(),"d")&" days"' in all cells where the out column is a date and the back column is blank
            $fixObj = New-Object PSObject |
                Add-Member NoteProperty WorksheetName $workShortName           -PassThru |
                Add-Member NoteProperty StartRow      $startRow                -PassThru |
                Add-Member NoteProperty RowCount      ($rowsUsed)              -PassThru
            $FixSet[$workShortName] = $fixObj
        }
        $startRow += $rowsUsed
    }
    Out-Null
}

#
# now apply overall formatting to the spreadsheet

    $ColumnInfo = @(
        @{ColumnField = "Name";                         ColumnTitle = "Name";                         ColumnWidth = 55; Wraps = $false; ColumnVisible = $true;  ColumnNumber = 1; NotFrozen = $false; ColumnUsed = $true;  ColumnSpare = ""; },
        @{ColumnField = "SonarID";                      ColumnTitle = "SonarID";                      ColumnWidth = 20; Wraps = $false; ColumnVisible = $true;  ColumnNumber = 2; NotFrozen = $false; ColumnUsed = $true;  ColumnSpare = ""; },
        @{ColumnField = "HomeURL";                      ColumnTitle = "HomeURL";                      ColumnWidth = 50; Wraps = $false; ColumnVisible = $true;  ColumnNumber = 3; NotFrozen = $false; ColumnUsed = $true;  ColumnSpare = ""; },
        @{ColumnField = "Subs";                         ColumnTitle = "Subs";                         ColumnWidth = 14; Wraps = $false; ColumnVisible = $true;  ColumnNumber = 4; NotFrozen = $false; ColumnUsed = $true;  ColumnSpare = ""; },
        @{ColumnField = "WorkBounds";                   ColumnTitle = "WorkBounds";                   ColumnWidth = 20; Wraps = $false; ColumnVisible = $true;  ColumnNumber = 4; NotFrozen = $false; ColumnUsed = $true;  ColumnSpare = ""; },
        @{ColumnField = "Agent";                        ColumnTitle = "Agent";                        ColumnWidth = 40; Wraps = $false; ColumnVisible = $true;  ColumnNumber = 4; NotFrozen = $false; ColumnUsed = $true;  ColumnSpare = ""; },
        @{ColumnField = "Genre";                        ColumnTitle = "Genre";                        ColumnWidth = 20; Wraps = $false; ColumnVisible = $true;  ColumnNumber = 4; NotFrozen = $false; ColumnUsed = $true;  ColumnSpare = ""; },
        @{ColumnField = "Out";                          ColumnTitle = "Out";                          ColumnWidth = 20; Wraps = $false; ColumnVisible = $true;  ColumnNumber = 4; NotFrozen = $false; ColumnUsed = $true;  ColumnSpare = ""; },
        @{ColumnField = "Back";                         ColumnTitle = "Back";                         ColumnWidth = 20; Wraps = $false; ColumnVisible = $true;  ColumnNumber = 4; NotFrozen = $false; ColumnUsed = $true;  ColumnSpare = ""; }
    )

    $excelPack = Open-ExcelPackage -Path $NewSubmissionFile
    $sheetIndex = 0
    Get-ExcelSheetInfo -Path $NewSubmissionFile | foreach {
        $sheetInfo = $_
        $work = $worksInOrder[$sheetIndex]
        if ($work.WorkName -like "Love*") {
            $work | Out-Null
        }
        $stats = New-Object PSObject |
            Add-Member NoteProperty Total      0 -PassThru |
            Add-Member NoteProperty Sales      0 -PassThru |
            Add-Member NoteProperty Rejections 0 -PassThru |
            Add-Member NoteProperty InProgress 0 -PassThru
        $SubmissionsHash.GetEnumerator() | % {$_.Value} | where { $_.WorkID -eq $work.WorkID } | foreach {
            $submissionXML = $_
            $stats.Total++
            if ($submissionXML.Sale -ne '0') {
                $stats.Sales++
            }
            elseif ($submissionXML.DateBack -notlike "1899*") {
                $stats.Rejections++
            }
            else {
                $stats.InProgress++
            }
        }
        $sheet = $excelPack.Workbook.Worksheets[$sheetInfo.Name]
        #for ($colIndex = 1; $colIndex -le $worksheetColumns.Length; $colIndex++) {
        #    $sheet.Column($colIndex) | Set-ExcelRange -VerticalAlignment Center 
        #}
        $ColumnInfo | where {$_.ColumnUsed} | foreach {
            $columnInfoRow = $_
            if (-not [string]::IsNullOrWhiteSpace($columnInfoRow.ColumnWidth)) {
                $sheet.Column($columnInfoRow.ColumnNumber) | Set-ExcelRange -Width $columnInfoRow.ColumnWidth
            }
           # if ($columnInfoRow.Wraps) {
           #     $sheet.Column($columnInfoRow.ColumnNumber) | Set-ExcelRange -WrapText
           # }
           # if ($columnInfoRow.ColumnVisible -eq $false) {
           #     $sheet.Column($columnInfoRow.ColumnNumber) | Set-ExcelRange -Hidden
           # }
        }
        #$sheet.Column(8) | Set-ExcelRange -WrapText -Width 75 # -Height 30 -VerticalAlignment Center 
        #$sheet.Column(1) | Set-ExcelRange -Width 40 # -Height 30 -VerticalAlignment Center 
        $sheet.Cells['A1'].Value = "Work Name"
        $sheet.Cells['B1'].Value = $work.WorkName

        $sheet.Cells['A2'].Value = "Work Length"
        $sheet.Cells['B2'].Value = $work.WorkLength

        $sheet.Cells['A3'].Value = "Sonar Work ID"
        $sheet.Cells['B3'].Value = $work.WorkID

        $sheet.Cells['A4'].Value = "Genre"
        $sheet.Cells['B4'].Value = $work.Genre

        $sheet.Cells['A5'].Value = "Sales"
        $sheet.Cells['B5'].Value = $stats.Sales

        $sheet.Cells['A6'].Value = "Active Submissions"
        $sheet.Cells['B6'].Value = $stats.InProgress
        #
        # let's handle the fixups
        $fixObj = $FixSet[$sheetInfo.Name]
        if ($fixObj -ne $null) {
            $fixObj | Out-Null
            for ($ix = 0; $ix -lt $fixObj.RowCount; $ix++) {
                $row = $fixObj.StartRow + $ix
                $outVal = $sheet.Cells["H$row"].Value
                if ($outVal -match "\d\d\d\d\-\d\d\-\d\d") {
                    $backVal = $sheet.Cells["I$row"].Value
                    if ([string]::IsNullOrWhiteSpace($backVal)) {
                        # need to put the formula '="out "&DATEDIF(H105,TODAY(),"d")&" days"' in all cells where the out column is a date and the back column is blank
                        $sheet.Cells["I$row"].Formula = '="out "&DATEDIF(H' + $row + ',TODAY(),"d")&" days"'
                    }
                }
            }
        }
        #
        # see https://github.com/dfinke/ImportExcel/issues/987
        if ($stats.InProgress -ne 0) {
            $sheet.TabColor = 'Red'
        }
        elseif ($stats.Sales -ne 0) {
            $sheet.TabColor = 'Green'
        }
        else {
            $sheet.TabColor = 'Yellow'
        }

        $sheet.Cells['A7'].Value = "Rejections"
        $sheet.Cells['B7'].Value = $stats.Rejections

        $sheet.Cells['A8'].Value = "Total submissions"
        $sheet.Cells['B8'].Value = $stats.Total

        $sheetIndex++
    }
    Close-ExcelPackage $excelPack

Write-Host "Done - acceptances:$script:Acceptances, rejections:$script:Rejections, on submission:$script:OnSubmission"
