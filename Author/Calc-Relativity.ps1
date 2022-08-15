#
# script to do basic match for near-light speed
#

param ([double]$AccG=1.0,
       [double]$FractionLightSpeed=0.6,
       [double]$TransitTimeHours = 1.0)

$LightSpeed = [double]299792458.0   # metres per second
$OneG = [double]9.8
$oneAU = 150.0 * 1000000.0 * 1000.0 # metres
$jupiterToMars = 3.7 * $oneAU # metres

$initialV = 0.0
$finalV = $LightSpeed * $FractionLightSpeed

$accel = $AccG * $OneG

#
# v = u + at
#
# so at = v - u
# t = (v-u)/a
#

$timeTakenSeconds = ($finalV - $initialV) / $accel

$oneYear = 24 * 3600 * 365.25

$yearsTaken = $timeTakenSeconds / $oneYear

Write-Host "time taken to accelerate to $FractionLightSpeed c at $AccG g = $yearsTaken years"

#
# now calculate the acceleration to travel half the distance mars to jupiter in 30 mins (accelerating to mid-point, decelerating to mid-point)

$halfDist = $jupiterToMars / 2.0
$halfTime = ($TransitTimeHours / 2.0) * 3600.0

#
# s = ut + 0.5at2
# so 2s = at2
# a = 2s/t2
#

$neededAccel = (2 * $halfDist) / ($halfTime * $halfTime)

$neededG = $neededAccel / $OneG

$maxVelocity = $neededAccel * $halfTime

$maxVelocityC = $maxVelocity / $LightSpeed

Write-Host "acceleration needed to travel from jupiter to mars in $TransitTimeHours hours = $neededG g, reaching a max velocity of $maxVelocityC c"
