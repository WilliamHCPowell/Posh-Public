#
# script to display projects geographically
# also work out what state the projects are in, so that they can be uploaded
# to Dynamics CRM
#

param ($infoFile="P:\Users\Bill\Documents\My Stories\Bobby's Dawn\LondonPoliceStations1830.csv",
       $WebFolder="K:\websites\hathiMaps\London1830"
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
function Get-ScriptDirectory
{
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
# see https://blogs.msdn.microsoft.com/koteshb/2010/02/12/powershell-how-to-find-details-of-operating-system/
#
$projectLocationFile          = "$sdir\Geospatial.js"
$BCCommandFile                = "$sdir\CopyFiles.bccommand"
#endregion

. ..\Postcode\LocationUtilities.ps1

$StationInfo = Import-Csv -LiteralPath $infoFile

$projData = $StationInfo | foreach { 
  $proj = $_
  if ($proj.Latitude -eq "") {
    #
    # try to find the location from the address information
    $result = Query-GeoApis -SearchString "$($proj.StreetAddress), London" -API "Google"
    $result | foreach {
      $proj.Latitude = $result.Latitude
      $proj.Longitude = $result.Longitude
      $proj.Source = $result.DataSource
    }
  }
  if ($proj.Latitude -ne "") {
#    if (([int]$proj.Opened -lt 1834) -and ([int]$proj.Closed -gt 1834)){
    if ($proj.Source -eq "Manual") {
      $pointColour = 'green'
    } else {
      $pointColour = 'magenta'
    }
    @{
       Title           = $proj.StationName;
       PointCount      = 0;
       Radius          = 20;
       CentreLatitude  = $proj.Latitude;
       CentreLongitude = $proj.Longitude;
       MinLatitude     = $proj.Latitude;
       MinLongitude    = $proj.Longitude;
       MaxLatitude     = $proj.Latitude;
       MaxLongitude    = $proj.Longitude;
       PointColour     = $pointColour;
       PopupText       = "<b>" + $proj.StationName + ":</b><br>" + $proj.StreetAddress + "<br>" + $proj.Division + " Division";
       Xcoord          = $proj.x;
       Ycoord          = $proj.y;
    }
  } else {
    Write-Host "No Location found for $($proj.StationName), $($proj.StreetAddress)"
  }
}

$StationInfo | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $infoFile

$point = @(51.522877,-0.142105,"<b>Trinity School</b><br>off Buckingham Street, where James Anderson met his end.")

..\SDS\Create-MapCode.ps1 -PointArray $projData `
                       -Style Marker `
                       -JSCodeFile $projectLocationFile `
                       -CentreOnLatitude $point[0] `
                       -CentreOnLongitude $point[1] `
                       -CentreMarkerHTML $point[2] `
                       -Scale 14

Copy-Item $projectLocationFile $WebFolder

$bcCommand =@"
# Load the base folders.
load create:right "profile:u67471010@bill-powell.co.uk?hathiMaps/BobbysDawn" "${WebFolder}"
filter "Geospatial.js"
# Copy different files left to right, delete orphans on right.
sync create-empty update:right->left
#
"@

$bcCommand | Set-Content $BCCommandFile

& "C:\Program Files (x86)\Beyond Compare 4\BCompare.exe" @$BCCommandFile  /closescript


Write-Host "Done"

