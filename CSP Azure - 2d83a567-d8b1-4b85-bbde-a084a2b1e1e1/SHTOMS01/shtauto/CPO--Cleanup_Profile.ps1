<#
Script til sletning af Chrome cache på Capto platformen
V1.0 by Jakob Strøm - 24-11-2016
#>

## Global variables
$cred = Get-AutomationPSCredential -Name "CPO--SVC_CPO_AzureAutomation"
Import-Module ActiveDirectory
$ErrorActionPreference = 'stop'

$debug              = $null
$Profiles           = '\\capto\data\uev_profiles'
$DC                 = 'CPO-AD-01.hosting.capto.dk'
$CustomerOU         = 'OU=Customer,OU=SYSTEMHOSTING,DC=hosting,DC=capto,DC=dk'
$Customers          = Get-ADOrganizationalUnit -SearchBase $CustomerOU -SearchScope OneLevel -Server $DC -Filter * -Credential $cred
$Users              = foreach($Name in $Customers.name){Get-ADUser -Filter "extensionAttribute1 -eq '$Name'" -Credential $cred}
$LogPath            = "\\capto\data\systemhosting\ProfileCleanup_LOG"
$logFile            = "Profile_Cleanup_" + (Get-Date -Format "dd-MM-yyyy_HH-mm-ss") + ".txt"
$PSDriveLogPath     = New-PSDrive V -PSProvider FileSystem -Root $LogPath -Credential $cred   ## Necessary for item get/delete because of Creds..
$PSDriveCustomer    = New-PSDrive Y -PSProvider FileSystem -Root $Profiles -Credential $cred  ## Necessary for item get/delete because of Creds..
$PSDriveLetter      = 'Y:\'
$PSDriveLogLetter   = 'V:\'

Write-Output ("PSdrives: " + (Get-PSDrive))  ## Info output to test PS drives

function Log([array]$text) {
	foreach ($txt in $text) {
            $txt | Out-File -FilePath ($PSDriveLogLetter + $logFile) -Encoding utf8 -Append
    }
}

Log ("Run started: " + (Get-Date -Format "dd-MM-yyyy_HH-mm-ss"))
Log ([string]($Users).Count + " users found. Starting compute")

foreach($i in $users){
    try{
        $UPMPath  = ($i.SamAccountName + ".V2")
        $UPMExist = Get-ChildItem ($PSDriveLetter + $UPMPath)
        }catch{
               Write-Output ("User " + $i.SamAccountName + " doesn't seem to have a UPM profile")
               if($debug){
               LOG ("User " + $i.SamAccountName + " doesn't seem to have a UPM profile")
               }
        }
    if($UPMExist){
        try{
            $Cache      = ($PSDriveLetter + $i.SamAccountName + ".v2\UPM_Profile\AppData\Local\Google\Chrome\User Data\Default\Cache")
            Get-ChildItem $Cache | Remove-Item -Recurse -Force
            Write-Output ("Users cache " + $i.SamAccountName + " deleted")
            Log ("Users cache " + $i.SamAccountName + " deleted")
            }catch{
               Write-Output ("Users cache " + $i.SamAccountName + " couldnt be deleted, got error: `n$_")
               if($debug){
               LOG ("Users cache " + $i.SamAccountName + " couldnt be deleted, got error: `n$_")
                    }
                }
        try{
            $MediaCache = ($PSDriveLetter + $i.SamAccountName + ".v2\UPM_Profile\AppData\Local\Google\Chrome\User Data\Default\Media Cache")
            Get-ChildItem $MediaCache | Remove-Item -Recurse -Force
            Write-Output ("Users Media Cache " + $i.SamAccountName + " deleted")
            Log ("Users Media Cache " + $i.SamAccountName + " deleted")
            }catch{
               Write-Output ("Users Media cache " + $i.SamAccountName + " couldnt be deleted, got error: `n$_")
               if($debug){
               LOG ("Users Media cache " + $i.SamAccountName + " couldnt be deleted, got error: `n$_")
                    }
                }
        try{
            $Temp       = ($PSDriveLetter + $i.SamAccountName + ".v2\UPM_Profile\AppData\Local\Temp")
            Get-ChildItem $Temp | Remove-Item -Recurse -Force
            Write-Output ("Users Temp " + $i.SamAccountName + " deleted")
            Log ("Users Temp " + $i.SamAccountName + " deleted")
            }catch{
               Write-Output ("Users Temp " + $i.SamAccountName + " couldnt be deleted, got error: `n$_")
               if($debug){
               LOG ("Users Temp " + $i.SamAccountName + " couldnt be deleted, got error: `n$_")
                    }
               }
        }
}
Log ("Run ended: " + (Get-Date -Format "dd-MM-yyyy_HH-mm-ss"))
Remove-PSDrive Y
Remove-PSDrive V