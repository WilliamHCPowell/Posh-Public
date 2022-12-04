#
# simulate methods of sorting a set of properties, good and bad, in a minimum number of presentations
#

param ($Stuff)

function Create-Distribution {
    [CmdletBinding()]
#    [OutputType([string])]
    param (
        [parameter (mandatory=$true)] [int]$TotalElements,
        [parameter (mandatory=$true)] [int]$BadElements
    )
    $AllElements = ('G' * $TotalElements) -split '' | where {-not [string]::IsNullOrWhiteSpace($_)}
    $CountOfB = 0
    while ($CountOfB -lt $BadElements) {
        $Index = Get-Random -Maximum $TotalElements -Minimum 0
        if ($AllElements[$Index] -eq 'G') {
            $AllElements[$Index] = 'B'
            $CountOfB++
        }
    }
    $AllElements
}

$script:Attempts = 0

function Test-Set {
    [CmdletBinding()]
#    [OutputType([string])]
    param (
        [parameter (mandatory=$true)] [array]$Set
    )
    $script:Attempts++
    if ('B' -in $Set) {
        throw "yuck!"
    }
    else {
        return $true
    }
}

function New-Range {
    [CmdletBinding()]
#    [OutputType([string])]
    param (
        [parameter (mandatory=$true)] [int]$Start,
        [parameter (mandatory=$true)] [int]$Size,
        [parameter (mandatory=$true)] [string]$Status
    )
    New-Object psobject | 
        Add-Member NoteProperty Start $Start -PassThru | 
        Add-Member NoteProperty Size $Size -PassThru | 
        Add-Member NoteProperty Status $Status -PassThru
}

function Get-LessEqualPowerOf2 {
    [CmdletBinding()]
#    [OutputType([string])]
    param (
        [parameter (mandatory=$true)] [int]$Target
    )
    $startAt = 64
    while ($startAt -lt $Target) {$startAt = $startAt * 2}
    while ($startAt -gt $Target) {$startAt = $startAt / 2}
    $startAt
}

#
# consider using a fibonacci series?

#
# powers of 2 fit
function Simulate-Distribution2 {
    [CmdletBinding()]
#    [OutputType([string])]
    param (
        [parameter (mandatory=$true)] [array]$Distribution
    )
    $script:Attempts = 0
    $rangeList = @()
    $range = New-Range 0 $Distribution.Length 'unknown'
    $rangeList += $range
    do {
        $firstUnknownRange = $rangeList | where {$_.Status -eq 'unknown'} | Select-Object -First 1
        $optimumsize = Get-LessEqualPowerOf2 $firstUnknownRange.Size
        if ($optimumsize -ne $firstUnknownRange.Size) {
            $excess = $firstUnknownRange.Size - $optimumsize
            $newstart = $firstUnknownRange.Start + $optimumsize
            $newRange = New-Range $newstart $excess $firstUnknownRange.Status
            $firstUnknownRange.Size = $optimumsize   # trim to power of 2
            $rangeList += $newRange
            $rangeList = $rangeList | Sort-Object -Property Start
        }
        #
        # let's see what we can discover about the current range
        try {
            #
            # we only ever test ranges that are a power of 2 in length
            $ItemsToTest = $Distribution[($firstUnknownRange.Start)..($firstUnknownRange.Start + $firstUnknownRange.Size - 1)]
            $val = Test-Set $ItemsToTest
            #
            # if we get here, the range is good
            $firstUnknownRange.Status = 'good'
        }
        catch {
            $e = $_
            if ($ItemsToTest.Count -eq 1) {
                $firstUnknownRange.Status = 'bad'
            }
            else {
                #
                # we need to split the array in 2 equal halves
                $optimumsize = $firstUnknownRange.Size / 2
                $excess = $firstUnknownRange.Size - $optimumsize
                $newstart = $firstUnknownRange.Start + $optimumsize
                $newRange = New-Range $newstart $excess $firstUnknownRange.Status
                $rangeList += $newRange
                $rangeList = $rangeList | Sort-Object -Property Start
                $firstUnknownRange.Size = $optimumsize
            }
        }
        $UnknownRangeCount = ($rangeList | where {$_.Status -eq 'unknown'} | Measure-Object).Count
    } while ($UnknownRangeCount -gt 0) 
    Write-Host -ForegroundColor Yellow "Algorithm 2 took $script:Attempts submissions to winnow $($Distribution -join '.')"
}

#
# binary chop
function Simulate-Distribution {
    [CmdletBinding()]
#    [OutputType([string])]
    param (
        [parameter (mandatory=$true)] [array]$Distribution
    )
    $script:Attempts = 0
    $UnknownQueue = New-Object System.Collections.Queue
    $SuspectQueue = New-Object System.Collections.Queue
    $SinBin = New-Object System.Collections.Queue
    $RightStuff = New-Object System.Collections.Queue
    for ($ix = 0; $ix -lt $Distribution.Length; $ix++) {
        $FieldItem = New-Object PSObject |
            Add-Member NoteProperty Status      'unknown'          -PassThru |
            Add-Member NoteProperty FieldName   $ix                -PassThru |
            Add-Member NoteProperty FieldValue  $Distribution[$ix] -PassThru
        $UnknownQueue.Enqueue($FieldItem)
    }
    while ($UnknownQueue.Count -gt 0) {
        $tempload = @()
        $pend = New-Object System.Collections.Queue
        #
        if ($true) {
            #
            # take items off the unknown queue
            $Len = [math]::Floor($UnknownQueue.Count / 2)
            $Len = [math]::Max(1,$Len)
            while ($Len -gt 0) {
                $FieldItem = $UnknownQueue.Dequeue()
                $Len--
                $tempload += $FieldItem.FieldValue
                $pend.Enqueue($FieldItem)
            }
        }
        if ($UnknownQueue.Count -eq 0) {
            #
            # take the rest of the unknown queue and move the suspect queue back to the unknown queue
            while ($SuspectQueue.Count -gt 0) {
                $FieldItem = $SuspectQueue.Dequeue()
                $UnknownQueue.Enqueue($FieldItem)
            }
        }
        try {
            $rsp = Test-Set $tempload
            #
            # if we get here, then everything is good and we can forget about the records we've processed
            while ($pend.Count -gt 0) {
                $FieldItem = $pend.Dequeue()
                $FieldItem.Status = 'Good'
                $RightStuff.Enqueue($FieldItem)
            }
        }
        catch {
            $f = $_
            $f | Out-Null
            #
            # if the count is 1, we put the record in the Sin Bin
            if ($pend.Count -eq 1) {
                $FieldItem = $pend.Dequeue()
                $FieldItem.Status = 'Bad'
                $SinBin.Enqueue($FieldItem)
            }
            else {
                while ($pend.Count -gt 0) {
                    $FieldItem = $pend.Dequeue()
                    $SuspectQueue.Enqueue($FieldItem)
                }
            }
        }
    }
    Write-Host -ForegroundColor Yellow "Algorithm 1 took $script:Attempts submissions to winnow $($Distribution -join '.')"
}

if ($Dist -eq $null) {
    $Dist = Create-Distribution -TotalElements 13 -BadElements 3
}

Write-Host -ForegroundColor Green "Distribution: $($Dist -join '.')"

Simulate-Distribution2 -Distribution $Dist

Simulate-Distribution -Distribution $Dist

if ($Dist2 -eq $null) {
    $Dist2 = Create-Distribution -TotalElements 97 -BadElements 3
}

Write-Host -ForegroundColor Green "Distribution: $($Dist2 -join '.')"

Simulate-Distribution2 -Distribution $Dist2

Simulate-Distribution -Distribution $Dist2

if ($Dist3 -eq $null) {
    $Dist3 = Create-Distribution -TotalElements 117 -BadElements 2
}

Write-Host -ForegroundColor Green "Distribution: $($Dist3 -join '.')"

Simulate-Distribution2 -Distribution $Dist3

Simulate-Distribution -Distribution $Dist3

if ($Dist4 -eq $null) {
    $Dist4 = Create-Distribution -TotalElements 131 -BadElements 3
}

Write-Host -ForegroundColor Green "Distribution: $($Dist4 -join '.')"

Simulate-Distribution2 -Distribution $Dist4

Simulate-Distribution -Distribution $Dist4
