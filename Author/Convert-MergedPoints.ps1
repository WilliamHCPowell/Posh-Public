$fieldList = "StationName,Category,Type,PointName,Latitude,Longitude,Coordinates,Confidence,x,y,Diff y,Diff Lat,Latitude per y pixel,Diff x,Diff Long,Longitude per x pixel,Calculated Origin Latitude,Calculated Origin Longitude,CalculatedLatitude,CalculatedLongitude,ErrorDistanceMetres,StreetAddress,Opened,Closed,Source,Division,DivisionName,Borough,Role,Capacity,Notes,URL,URL2,Book,Page" -split ','
$x = Import-Csv .\1830PointsMerged.csv | foreach {
    $point = $_
    foreach ($fieldName in $fieldList) {
        if (-not [string]::IsNullOrWhiteSpace($point.$fieldName)) {
            switch -Wildcard ($fieldName) {
                "*latitude*" {
                        $point.$fieldName = [double]$point.$fieldName
                    }
                "*longitude*" {
                        $point.$fieldName = [double]$point.$fieldName
                    }
                "*diff*" {
                        $point.$fieldName = [double]$point.$fieldName
                    }
                "Confidence" {
                        $point.$fieldName = [int]$point.$fieldName
                    }
                "ErrorDistanceMetres" {
                        $point.$fieldName = [double]$point.$fieldName
                    }
 #               "Opened" {
 #                       $point.$fieldName = [int]$point.$fieldName
 #                   }
 #               "Closed" {
 #                       $point.$fieldName = [int]$point.$fieldName
 #                   }
            }
        }
    }
    $point
}
$x | ConvertTo-Json > .\1830PointsMerged.json