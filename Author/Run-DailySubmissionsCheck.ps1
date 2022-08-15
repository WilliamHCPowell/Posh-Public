#
# script to run daily submissions check
#

param ($JobName = "ManualSubmissionsCheck")

cd C:\Projects\Author
$sdir = (Get-Location).Path

Out-File ".\${JobName}_Started.timestamp"

.\Test-Submissions.ps1 -JobName $JobName

Out-File ".\${JobName}_Stage1Complete.timestamp"

.\Analyse-Sonar.ps1 -JobName $JobName

Out-File ".\${JobName}_Stage2Complete.timestamp"

.\Plot-Submissions.ps1 -JobName $JobName

Out-File ".\${JobName}_Finished.timestamp"

