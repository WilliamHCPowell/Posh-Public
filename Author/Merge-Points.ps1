#
# script to merge the two different sources of geographical information
#

param ($infoFile="P:\Users\Bill\Documents\My Stories\Bobby's Dawn\LondonPoliceStations1830.csv",
       $calibrationFile="P:\Users\Bill\Documents\My Stories\Bobby's Dawn\CalibrationPoints.xlsx"
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
$AllPointsMergedFile          = "$sdir\1830PointsMerged.csv"
$AllPointsMergedJson          = "$sdir\1830PointsMerged.json"
$WebFolder                    = "$sdir\Upload"
#endregion

$HQ = @("St Pauls Cathedral",51.513812,-0.097246)

. ..\Postcode\LocationUtilities.ps1

$StationInfo = Import-Csv -LiteralPath $infoFile

$StationCache = @{}

$StationInfo | where {[string]$_.StationName -ne ""} | foreach {
  $StationCache[$_.StationName] = $_
}

$AllPointInfo = $(.\Get-CalibrationPoints.ps1)

$AllPointCache = @{}

$AllPointInfo | where {[string]$_.StationName -ne ""} | foreach {
  $AllPointCache[$_.StationName] = $_
}

#
# work out common and unique fields
$stationFields = $StationInfo[0].PSObject.properties | foreach { $_.Name }

$allPointFields = $AllPointInfo[0].PSObject.properties | foreach { $_.Name }

$commonFields = $stationFields | where {$_ -cin $allPointFields}

$stationOnlyFields = $stationFields | where {$_ -cnotin $commonFields}

$calibrationOnlyFields = $allPointFields | where {$_ -cnotin $commonFields}

#
# check that every entry in $StationInfo is also in $AllPointCache
$fatal = $false
$StationInfo | foreach {
  $station = $_
  $calibrationPoint = $AllPointCache[$station.StationName]
  if ($calibrationPoint -eq $null) {
    Write-Host @errorColours "'$($station.StationName)' not found in Calibration Point list"
    $fatal = $true
  } else {
    $commonFields | foreach {
      $fieldName = $_
      if ($station.$fieldName -ne $calibrationPoint.$fieldName) {
        Write-Host @warningColours "($($station.StationName))->$fieldName : Station = '$($station.$fieldName)', Calibration = '$($calibrationPoint.$fieldName)'"
        if ($station.$fieldName -ne "") {
          $calibrationPoint.$fieldName = $station.$fieldName
        }
      } else {
        # Write-Host @infoColours "($($station.StationName))->$fieldName : Station = '$($station.$fieldName)', Calibration = '$($calibrationPoint.$fieldName)'"
      }
    }
  }
}

if ($fatal) {
  Write-Host @errorColours "Fatal error encountered - exit"
  exit 0
}

$point = @{}

$calibrationPointInfo = $AllPointInfo | where {$_.Confidence -eq 0} | foreach {
  $point[$_.PointName] = $_
  $_
}

#
# for the mathematics behind this calculation, see the document "Calculating vector - corrected.pdf" in this folder
# which describes how to map a point D, referenced by 3 calibration points A, B and C, defined in two spaces
# from the coordinates in one space to the corrdinates in the other space, by expressing the vector AD as 
#   AD = s*AB + t*BC 
# Knowing D in one space and A, B and C in both spaces, (and assuming uniform rotation/deformation) we can 
# calculate s and t in one coordinate space, from which we can calculate AD in the other space, and thus D.
#
$AllPointsMerged = $AllPointInfo | foreach {
  $pointD = $_ | 
    Add-Member NoteProperty CalculatedLatitude  "" -PassThru | 
    Add-Member NoteProperty CalculatedLongitude "" -PassThru | 
    Add-Member NoteProperty ErrorDistanceMetres "" -PassThru
  if (($pointD.Confidence -ne 0) -and ($pointD.StationName -ne "")) {
    $pointC = $point["C"]
    $pointB = $point["B"]
    $pointA = $point["A"]
    if ([string]$pointD.x -eq "") {
      # note 
      $xa = [double]$pointA.Longitude
      $ya = [double]$pointA.Latitude
      $xb = [double]$pointB.Longitude
      $yb = [double]$pointB.Latitude
      $xc = [double]$pointC.Longitude
      $yc = [double]$pointC.Latitude
      $xd = [double]$pointD.Longitude
      $yd = [double]$pointD.Latitude
      $t = ((($yd - $ya) * ($xb - $xa)) - (($xd - $xa) * ($yb - $ya))) / 
           ((($yc - $yb) * ($xb - $xa)) - (($xc - $xb) * ($yb - $ya)))
      $s = ($xd - $xa) - ($t * ($xc - $xb)) / ($xb - $xa)
      $resxa = [double]$pointA.x
      $resya = [double]$pointA.y
      $resxb = [double]$pointB.x
      $resyb = [double]$pointB.y
      $resxc = [double]$pointC.x
      $resyc = [double]$pointC.y
      $resxd = ($s * ($resxb - $resxa)) + ($t * ($resxc - $resxb))
      $resyd = ($s * ($resyb - $resya)) + ($t * ($resyc - $resyb))
      $pointD.y = $resya+$resyd
      $pointD.x = $resxa+$resxd
      Write-Host @infoColours "$($pointD.StationName) is at $($pointD.x), $($pointD.y)"
    } else { # if ([string]$pointD.Longitude -eq "") {
      # note 
      $xa = [double]$pointA.x
      $ya = [double]$pointA.y
      $xb = [double]$pointB.x
      $yb = [double]$pointB.y
      $xc = [double]$pointC.x
      $yc = [double]$pointC.y
      $xd = [double]$pointD.x
      $yd = [double]$pointD.y
      $t = ((($yd - $ya) * ($xb - $xa)) - (($xd - $xa) * ($yb - $ya))) / 
           ((($yc - $yb) * ($xb - $xa)) - (($xc - $xb) * ($yb - $ya)))
      $s = (($xd - $xa) - ($t * ($xc - $xb))) / ($xb - $xa)
      $resxa = [double]$pointA.Longitude
      $resya = [double]$pointA.Latitude
      $resxb = [double]$pointB.Longitude
      $resyb = [double]$pointB.Latitude
      $resxc = [double]$pointC.Longitude
      $resyc = [double]$pointC.Latitude
      $resxd = ($s * ($resxb - $resxa)) + ($t * ($resxc - $resxb))
      $resyd = ($s * ($resyb - $resya)) + ($t * ($resyc - $resyb))
      $pointD.CalculatedLatitude = $resya+$resyd
      $pointD.CalculatedLongitude = $resxa+$resxd
      #
      # calculate the distance from the Lat,Long if known
      if ($pointD.Latitude -ne "") {
        $diffLat = 250.0 * ($pointD.Latitude - $pointD.CalculatedLatitude) /$DegreesLatitudePer250m
        $diffLong = 250.0 * ($pointD.Longitude - $pointD.CalculatedLongitude) / $DegreesLongitudePer250m
        $distance = [math]::sqrt(($difflat * $diffLat) + ($diffLong * $diffLong))
        $pointD.ErrorDistanceMetres = $distance
      }
      Write-Host @infoColours "$($pointD.StationName) is at $($pointD.CalculatedLatitude), $($pointD.CalculatedLongitude)"
    }
  }
  $pointD
}

$bigPoints = $AllPointsMerged | where {$_.StationName -ne ""} | foreach {
  $point = $_
  $station = $StationCache[$point.StationName]
  $stationOnlyFields | foreach {
    $fieldName = $_
    if ($station -eq $null) {
      $val = $null
    } else {
      $val = $station.$fieldName
    }
    $point = $point | Add-Member NoteProperty -Name $fieldName -Value $val -PassThru
  }
  $point
}

$bigPoints | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $AllPointsMergedFile
$bigPoints | ConvertTo-Json > $AllPointsMergedJson

Copy-Item $AllPointsMergedJson $WebFolder

$bcCommand =@"
# Load the base folders.
load create:right "profile:u67471010@bill-powell.co.uk?hathiMaps" "${WebFolder}"
filter "*.json"
# Copy different files right to left.
sync create-empty update:right->left
#
"@

$bcCommand | Set-Content $BCCommandFile

& "C:\Program Files (x86)\Beyond Compare 4\BCompare.exe" @$BCCommandFile  /closescript


Write-Host "Done"

