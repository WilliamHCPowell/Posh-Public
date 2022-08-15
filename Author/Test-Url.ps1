#
# script to handle the testing of submission URLs - separated out so that it can be called from a script block
#

param ([string]$URL,
       [switch]$UseInvokeWR)

function Read-HtmlPage {
    param ([Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][String] $Uri)

    # Invoke-WebRequest and Invoke-RestMethod can't work properly with UTF-8 Response so we need to do things this way.
    [Net.HttpWebRequest]$WebRequest = [Net.WebRequest]::Create($Uri)
    $WebRequest.Timeout = 15 * 1000  # milliseconds
    [Net.HttpWebResponse]$WebResponse = $WebRequest.GetResponse()
    $Reader = New-Object IO.StreamReader($WebResponse.GetResponseStream())
    $Response = $Reader.ReadToEnd()
    $Reader.Close()

    # Create the document class
    [mshtml.HTMLDocumentClass] $Doc = New-Object -com "HTMLFILE"
    $Doc.IHTMLDocument2_write($Response)
    
    # Returns a HTMLDocumentClass instance just like Invoke-WebRequest ParsedHtml
    $Doc
}

function Test-URL-Invoke-WebRequest ([string]$URL) {
  $page = Invoke-WebRequest $URL -UseDefaultCredentials -TimeoutSec 15
  return $page.ParsedHtml.body.outerText
}

function Test-URL-Read-HtmlPage ([string]$URL) {
  $page = Read-HtmlPage -Uri $URL
  $unparsedText = $page.all | foreach { $_.outerText }
  $unparsedText = $unparsedText -join "`n"
  return $unparsedText
}

  if ($URL -eq "") {
    return "Blank"  
  } elseif ($URL.StartsWith("mailto:")) {
    return "Mail"  
  } else {
    $retVal = "Open"
  }

  try {
    if ($UseInvokeWR) {
      $textToCheck = Test-URL-Invoke-WebRequest -URL $URL
    } else {
      $textToCheck = Test-URL-Read-HtmlPage -URL $URL
    }
  }
  catch {
    $e = $_
    Write-Host @errorColours "Error : $e"
    Write-Host @errorColours "URL : $URL"
    switch -Wildcard ($e) {
      "*404*" {
          $retVal = "Not Found"
        }
      "*Not Found*" {
          $retVal = "Not Found"
        }
      "*503*" {
          $retVal = "Busy"
        }
      default {
          $retVal = "Fault"
        }
    }
  }

  if ($textToCheck -ne $null) {
    #
    # for the sake of brevity, limit the length of text to be checked
    $textToCheck = $textToCheck.Substring(0,[math]::Min($textToCheck.Length, 4096))
    switch -regex ($textToCheck) {
      ".*Closed.*" {
          $retVal = "Closed"
        }
      ".*not accepting.*" {
          $retVal = "Closed"
        }
      ".*no open calls.*" {
          $retVal = "Closed"
        }
      ".*on hiatus.*" {
          $retVal = "Closed"
        }
    }
  }


  $retVal
