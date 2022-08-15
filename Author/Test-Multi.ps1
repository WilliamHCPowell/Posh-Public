#
# script to test async operation of a script block to check URLs
#

param ($nThreads=2,
       [switch]$UseInvokeWR
      )

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
function Get-ScriptDirectory
{
  $Invocation = (Get-Variable MyInvocation -Scope 1).Value
  Split-Path $Invocation.MyCommand.Path
}

$sdir = Get-ScriptDirectory
#endregion

#region Display Colours
#
# nice colours when using Write-Host
#
$infoColours    = @{foreground="cyan"}
$warningColours = @{foreground="yellow"}
$errorColours   = @{foreground="red"}
$debugColours   = @{foreground="green"}

#endregion

#region File Management
#
# Having identified the current working directory, we can now set up paths for the
# various cache files and report files used by the script.
#
$TestDataFile          = $sdir + "\" + "TestData.csv"                  #
$webScriptFile         = $sdir + "\" + "Test-Url.ps1"                  #

#endregion

$urlsToTest = Import-Csv -LiteralPath $TestDataFile | Select-Object -First 20

cls

.\Test-Async.ps1 -nThreads $nThreads -runOption Sync -urlsToTest $urlsToTest -UseInvokeWR: $UseInvokeWR
.\Test-Async.ps1 -nThreads $nThreads -runOption Job -urlsToTest $urlsToTest -UseInvokeWR: $UseInvokeWR
.\Test-Async.ps1 -nThreads $nThreads -runOption RunSpace -urlsToTest $urlsToTest -UseInvokeWR: $UseInvokeWR
