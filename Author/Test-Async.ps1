#
# script to test async operation of a script block to check URLs
#

param ($nThreads=2,
      [ValidateSet('Sync','Job','RunSpace')]
       $runOption="Job",
       $urlsToTest,
       [switch]$UseInvokeWR
      )

#  cls

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

# This script block will test a URL
$ScriptBlock = {
   Param (
      $Context
   )
   # Write-Host $Context.SubmitUrl
   $res = C:\Projects\Author\Test-URL $Context.SubmitUrl -UseInvokeWR: $UseInvokeWR
#   $res = Invoke-Expression "$webScriptFile -URL '$($Context.SubmitUrl)'"
   $Context["Result"] = $res
   $Context
}

if ($urlsToTest -eq $null) {
  $urlsToTest = Import-Csv -LiteralPath $TestDataFile
}

$numThreads = $nThreads
Write-Host "checking $($urlsToTest.Length) URLs using $runOption with $numThreads background threads"

if ($runOption -eq "RunSpace") {
  # Create session state
  $myString = "this is session state!"
  $sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
  # $sessionstate.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "myString" ,$myString, "example string"))
   
  # Create runspace pool consisting of $numThreads runspaces
  $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $numThreads, $sessionState, $Host)
  $RunspacePool.Open()
}

$urlsToTest | foreach {
  $ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()
  $info = $_
  $Context = @{Title=$info.Title;ID=$info.IG;SubmitURL=$info.SubmitURL}
  switch ($runOption) {
    "Sync" {
        $OutputContext = Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Context
      }
    "Job" {
        $job = Start-Job  -ScriptBlock $ScriptBlock -ArgumentList $Context

        $OutputContext = $job | Wait-Job | Receive-Job
      }
    "RunSpace" {
        $rJob = [powershell]::Create().AddScript($ScriptBlock).AddParameter("Context",$Context)
        $rJob.RunspacePool = $RunspacePool
        $Handle = $rJob.BeginInvoke()
        # Write-Host "Started"
        while ($Handle.IsCompleted -eq $false) {
          Start-Sleep -Seconds 1
          # Write-Host -NoNewline "." 
        }
        $OutputContext = $rJob.EndInvoke($Handle)
      }
  }
  Write-Host @debugColours "$($OutputContext.SubmitURL) $($OutputContext.Result) $($ElapsedTime.Elapsed.TotalSeconds) seconds"
}

if ($runOption -eq "RunSpace") {
  $RunspacePool.Dispose()
}

reportPhaseComplete "checked $($urlsToTest.Length) URLs using $runOption"

exit 0
#
# use jobs
#
$urlsToTest | foreach {
  $info = $_
  $Context = @{Title=$info.Title;ID=$info.IG;SubmitURL=$info.SubmitURL}
  Write-Host "$($OutputContext.SubmitURL) $($OutputContext.Result)"
}

exit 0
Write-Host ""
$numThreads = 5
Write-Host "Now lets try creating 50 files by running up $numThreads background threads"

# Create session state
$myString = "this is session state!"
$sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$sessionstate.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "myString" ,$myString, "example string"))
   
# Create runspace pool consisting of $numThreads runspaces
$RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $numThreads, $sessionState, $Host)
$RunspacePool.Open()

$startTime = Get-Date
$Jobs = @()
$urlsToTest | % {
  $info = $_
  $Context = @{Title=$info.Title;ID=$info.IG;SubmitURL=$info.SubmitURL}

    $Job = [powershell]::Create().AddScript($ScriptBlock).AddParameter("Context",$Context)
    $Job.RunspacePool = $RunspacePool
    $Jobs += New-Object PSObject -Property @{
      RunNum = $_
      Job = $Job
      Result = $Job.BeginInvoke()
   }
}
 
Write-Host "Waiting.." -NoNewline
Do {
   Write-Host "." -NoNewline
   Start-Sleep -Seconds 1
} While ( $Jobs.Result.IsCompleted -contains $false) #Jobs.Result is a collection

$endTime = Get-Date
$totalSeconds = "{0:N4}" -f ($endTime-$startTime).TotalSeconds
Write-Host "All files created in $totalSeconds seconds"


