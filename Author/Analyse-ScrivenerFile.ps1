#
# script to count words and chapters in a scrivener project
#

param ($RTFFilePath = "P:\IDrive-Sync\My Reviews\Laura Sorrese\Past Is Prologue.scriv\Files\Data\655A407E-0542-48B9-8A52-1F0586D129AE\content.rtf",
       $ScrivenerFile = "P:\IDrive-Sync\My Reviews\Laura Sorrese\Past Is Prologue.scriv\Past Is Prologue.scrivx")

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

Import-Module ImportExcel

$word = New-Object -ComObject word.application
$word.visible = $false

$SaveChanges = $false

function Count-WordsInDoc ($RTFFilePath,$ChapterName) {
    $WordDoc = $RTFFilePath

    $doc = $word.documents.open($WordDoc)
    $totalParagraphs = $doc.Paragraphs.Count
    $itemCount = 0
    $wordCount = 0
    $activity = "Counting Words in chapter $ChapterName"
    $doc.Paragraphs | Select-Object -Skip 0 | foreach {
        $paragraph = $_
        $wordList = $paragraph.Range.Words | foreach {$_.Text}                                                                                                                                                                                                    
        $wl2 = $wordList | where {$_ -match ".*[0-9A-Z]+.*"}                                                                                                                                                                                                      
        $wordCount += $wl2.Count
        $itemCount++
        $status = "Analysed $itemCount of $totalParagraphs paragraphs"
        Write-Progress -activity $activity -status $status -PercentComplete (($itemCount / $totalParagraphs) * 100)
    }

    $doc.close([ref]$SaveChanges)
    Write-Progress -activity $activity -Completed -Status "Success"
    return $wordCount
}

function New-AnalysisChallenge {
    New-Object PSObject |
        Add-Member NoteProperty Analysis  "" -PassThru |
        Add-Member NoteProperty Challenge "" -PassThru
}

function Extract-AnalysisChallenge ($RTFFilePath) {
    $WordDoc = $RTFFilePath

    $doc = $word.documents.open($WordDoc)
    $totalParagraphs = $doc.Paragraphs.Count
    $itemCount = 0
    $wordCount = 0
    $activity = "Counting Words"
    $Target = "None"
    $CachedText = @{}
    $doc.Paragraphs | Select-Object -Skip 0 | foreach {
        $paragraph = $_
        $Text = $paragraph.Range.Text -replace "[`r`n\s]+$",''
        switch ($Text) {
            "Analysis" {$Target = $_}
            "Challenge" {$Target = $_}
            default {
                    $Previous = $CachedText[$Target]
                    if ($Previous -eq $null) {
                        $Previous = @($_)
                    }
                    else {
                        $Previous += $_
                    }
                    $CachedText[$Target] = $Previous
                }
        }
    }

    $doc.close([ref]$SaveChanges)
    $obj = New-AnalysisChallenge
    $obj.Analysis = $CachedText["Analysis"] -join "`r`n"
    $obj.Challenge = $CachedText["Challenge"] -join "`r`n"
    $obj
}

$Text = Get-Content -LiteralPath $ScrivenerFile
$ScrivenerXMLDoc = [xml]$Text

$ProjectItem = Get-Item $ScrivenerFile
$ProjectRoot = $ProjectItem.DirectoryName
$FilesFolder = Join-Path $ProjectRoot "Files"
$DataFolder = Join-Path $FilesFolder "Data"

$bookWordCount = 0

$ChapterLookup = @{}

$ChapterNumber = 0

$ActName = ""

$DataCollection = $ScrivenerXMLDoc.ScrivenerProject.Binder.BinderItem | foreach {
    $BinderItem = $_
    if ($BinderItem.Title -eq "Manuscript") {
        #
        # iterate through children
        $BinderItem.Children.BinderItem | foreach {
            #
            $ChapterFolder = $_
            $ChapterDataFolder = Join-Path $DataFolder $ChapterFolder.UUID
            $SynopsisPath = Join-Path $ChapterDataFolder "synopsis.txt"
            if (Test-Path -LiteralPath $SynopsisPath) {
                $SynopsisTxt = (Get-Content -LiteralPath $SynopsisPath) -join "`r`n"
            }
            $NotesPath = Join-Path $ChapterDataFolder "notes.rtf"
            if (Test-Path -LiteralPath $NotesPath) {
                $AnalysisObj = Extract-AnalysisChallenge $NotesPath
            }
            else {
                $AnalysisObj = New-AnalysisChallenge
            }
            $ChapterName = ""

            $ChapterNumber++
            $ChapterWordCount = @()
            $SceneCollection = @($ChapterFolder.Children.BinderItem)
            $SceneNumber = 1
            $SceneCollection | foreach {
                $Scene = $_
                if ($Scene.Title -eq "Prologue") {
                    $ChapterNumber = 0
                    $ChapterName = $Scene.Title
                }
                elseif ($Scene.Title -match "^Chapter (\d+)\:\s*(\S.*\S)\s*$") {
                    $ChapterNumber = [int]($Matches[1])
                    $ChapterName = $Matches[2]
                }
                elseif ($Scene.Title -match "^ACT\s*(\S.*\S)\s*$") {
                    $ActName = "ACT " + $Matches[1]
                    $ChapterName = $Matches[2]
                }
                if ($Scene.MetaData.IncludeInCompile -eq "Yes") {
                    if ($Scene.MetaData.CustomMetaData -ne $null) {
                        $customMetaData = $Scene.MetaData.CustomMetaData
                        $customMetaData.MetaDataItem | foreach {
                            if ($_.FieldId -eq "intensity") {
                                $Intensity = $_.Value
                            }
                        }
                    }
                    else {
                        $Intensity = 0
                    }
                    $SceneFolder = $Scene.UUID
                    $SceneFolderPath = Join-Path $DataFolder $SceneFolder
                    Get-ChildItem $SceneFolderPath -Include "*.$($Scene.MetaData.FileExtension)" -Recurse | foreach {
                        $ChapterTextFileObj = $_
                        $wordCount = Count-WordsInDoc -RTFFilePath $ChapterTextFileObj.FullName `
                                                      -ChapterName $ChapterName
                        $ChapterWordCount += $wordCount
                        $bookWordCount += $wordCount
                        New-Object PSObject |
                            Add-Member NoteProperty Act               $ActName               -PassThru |
                            Add-Member NoteProperty ChapterNumber     $ChapterNumber         -PassThru |
                            Add-Member NoteProperty ChapterName       $ChapterName           -PassThru |
                            Add-Member NoteProperty SceneNumber       $SceneNumber           -PassThru |
                            Add-Member NoteProperty SceneName         $Scene.Title           -PassThru |
                            Add-Member NoteProperty Intensity         $Intensity             -PassThru |
                            Add-Member NoteProperty WordCount         $wordCount             -PassThru |
                            Add-Member NoteProperty Synopsis          $SynopsisTxt           -PassThru |
                            Add-Member NoteProperty Analysis          $AnalysisObj.Analysis  -PassThru |
                            Add-Member NoteProperty Challenge         $AnalysisObj.Challenge -PassThru
                        $SceneNumber++
                        $SynopsisTxt = ""
                        $AnalysisObj = New-AnalysisChallenge
                    }
                }
            }
            Write-Host @infoColours "Chapter $ChapterNumber : '$ChapterName' has $($ChapterWordCount.Count) scenes with $($ChapterWordCount -join '|') words"
        }
    }
}

$TargetCsv = Join-Path $ProjectRoot "..\$($ProjectItem.BaseName).csv"
$TargetXlsx = Join-Path $ProjectRoot "..\$($ProjectItem.BaseName).xlsx"

$SceneColors = "Red,Green,Blue,Yellow,Magenta,Cyan" -split ','
$Index = 0
$DataCollection | 
    Select-Object -Property Act,ChapterNumber,ChapterName,SceneNumber,SceneName,SceneLabel,SceneColor,Intensity,WordCount,Synopsis,Analysis,Challenge | 
    foreach {
        $_.SceneLabel = "$($_.ChapterNumber).$($_.SceneNumber)"; 
        $_.SceneColor = $SceneColors[$Index]
        $Index++
        $Index = $Index % $SceneColors.Count
        $_
    } | 
    Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $TargetCsv

$DataCollection | Export-Excel -Path $TargetXlsx `
                               -AutoSize `
                               -AutoFilter `
                               -WorksheetName $ProjectItem.BaseName `
                               -StartRow 1 `
                               -StartColumn 1 `
                               -FreezePane 3,8 `
                               -TableName $($ProjectItem.BaseName -replace "\s+",'') `
                               -Title "Analysis: $($ProjectItem.BaseName)"

$excelPack = Open-ExcelPackage -Path $TargetXlsx
$sheet = $excelPack.Workbook.Worksheets[$ProjectItem.BaseName]

1..10 | % {$sheet.Column($_)  | Set-ExcelRange -VerticalAlignment Top -HorizontalAlignment Left}
$sheet.Column(5)  | Set-ExcelRange -Width 32 -WrapText -VerticalAlignment Top
$sheet.Column(8)  | Set-ExcelRange -Width 75 -WrapText -VerticalAlignment Top
$sheet.Column(9)  | Set-ExcelRange -Width 75 -WrapText -VerticalAlignment Top
$sheet.Column(10) | Set-ExcelRange -Width 75 -WrapText -VerticalAlignment Top

Close-ExcelPackage $excelPack

#
# close down the automation and release resources
#
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Remove-Variable -Name word
[gc]::collect()
[gc]::WaitForPendingFinalizers()

Write-Host "Done - $bookWordCount words counted"
