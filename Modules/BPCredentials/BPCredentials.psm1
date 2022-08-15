#
# functions, etc
#

$script:ConfigSecretFolder = "C:\Projects\Config"
$script:effectiveUser = $env:USERNAME
$script:CredentialRepository = Join-Path $script:ConfigSecretFolder "Credentials_${effectiveUser}.xml"

$script:CredentialTemplate = @"
{
    "CredentialUniqueName":  "",
    "ValidFrom":  "",
    "ValidTo":  "",
    "URL":  "",
    "AccountName":  "",
    "Secret":  "",
    "Credential":  {
                       "UserName":  "",
                       "Password":  ""
                   }
}
"@

function New-ClearCredential {
    $script:CredentialTemplate | ConvertFrom-Json
<#
    $cred = New-Object PSObject |
       Add-Member NoteProperty UserName "" -PassThru |
       Add-Member NoteProperty Password "" -PassThru
    New-Object PSObject |
       Add-Member NoteProperty CredentialUniqueName "" -PassThru |
       Add-Member NoteProperty ValidFrom "" -PassThru |
       Add-Member NoteProperty ValidTo "" -PassThru |
       Add-Member NoteProperty URL "" -PassThru |
       Add-Member NoteProperty AccountName "" -PassThru |
       Add-Member NoteProperty Secret "" -PassThru |
       Add-Member NoteProperty Credential $cred -PassThru
#>
}

function Restore-AllCredentials {
    if (Test-Path -LiteralPath $script:CredentialRepository) {
        Import-Clixml -LiteralPath $script:CredentialRepository
    }
    else {
        @{}
    }
}

function Save-ClearCredential ($ClearCredential) {
    $AllCredentials = Restore-AllCredentials

    $secpasswd = ConvertTo-SecureString $ClearCredential.Secret -AsPlainText -Force
    $ClearCredential.Credential = New-Object System.Management.Automation.PSCredential ($ClearCredential.AccountName, $secpasswd)
    $ClearCredential.Secret = ""

    $AllCredentials[$ClearCredential.CredentialUniqueName] = $ClearCredential

    $AllCredentials | Export-Clixml -LiteralPath $script:CredentialRepository
}

function Restore-ClearCredential ($CredentialUniqueName) {
    $AllCredentials = Restore-AllCredentials

    $ClearCredential = $AllCredentials[$CredentialUniqueName]
    $ClearCredential.Secret = $ClearCredential.Credential.GetNetworkCredential().Password

    $ClearCredential
}

function Get-ClearCredentialUniqueNames {
    $AllCredentials = Restore-AllCredentials
    $AllCredentials.GetEnumerator() | foreach {$_.Key} | Sort-Object
}

Export-ModuleMember New-ClearCredential,Save-ClearCredential,Restore-ClearCredential,Get-ClearCredentialUniqueNames
