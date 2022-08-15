#
# script to remove unwanted metadata from Scrivener-generated Word Documents
#
#
# see https://docs.microsoft.com/en-us/office/vba/api/word.range.select and related pages for 
# manipulation of text
#

param ($WordDoc = 'P:\Users\Bill\Documents\My Stories\Bobby''s Dawn\Bobby''s Dawn - Beta - 3rd Draft final post CBC.docx')

#
# set up the automation
#
$word = New-Object -ComObject word.application
$word.visible = $false

$SaveChanges = $true

    $doc = $word.documents.open($WordDoc)
    $totalFiles = $doc.Paragraphs.Count
    $itemCount = 0
    $activity = "Removing metadata"
    $doc.Paragraphs | foreach {
        $paragraph = $_
        if ($paragraph.Range.Text -match "^ChapterDate\:.*$") {
            # we want to replace the text
            $paragraph | Out-Null
            $paragraph.Range.Select()
            $selection = $doc.ActiveWindow.Selection
            $newText = $selection.Text -replace "^\S+\s+",''
            $selection.Text = $newText
        }
        elseif ($paragraph.Range.Text -match "^(Created|Modified|Status|Label|Event Start Date)\:.*$") {
            #
            # we want to delete this
            $paragraph.Range.Select()
            $paragraph.Range.Delete() | Out-Null
        }
        else {
            $paragraph | Out-Null
            $paragraph.Range.Select()
        }
        $itemCount++
        $status = "Analysed $itemCount of $totalFiles items"
        Write-Progress -activity $activity -status $status -PercentComplete (($itemCount / $totalFiles) * 100)
    }

    $doc.close([ref]$SaveChanges)
Write-Progress -activity $activity -Completed -Status "Success"

#
# close dow the automation and release resources
#
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Remove-Variable -Name word
[gc]::collect()
[gc]::WaitForPendingFinalizers()

Write-Host "Done"
