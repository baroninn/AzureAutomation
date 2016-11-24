<#
    .DESCRIPTION
    Imports SH contacts into all user mailboxes..   

    .NOTES
        AUTHOR: Jakob Strøm
        LASTEDIT: September 14, 2016


#>

$logFile  = "C:\Scripts\Logs\Contactsimport_" + (Get-Date -Format "dd-MM-yyyy_HH-mm-ss") + ".txt"

function Log([array]$text) {
	foreach ($txt in $text) {
		
		$txt | Out-File -FilePath $logFile -Encoding "UTF8" -Append
	}
}

$ErrorActionPreference = 'stop'

$cred = Get-AutomationPSCredential -Name "EXCHANGE--SVC_AzureAutomation"
$Contacts = "\\exch023C1CAS\c$\contacts.pst"
# $VerbosePreference='Continue'

#  Get Organization, only used when run manually.
#$Organisation = Read-Host "Customer initials.. (Leave blank for full report)" 

# make the connection to exchange:
function Connect-exchange{
    param(
         [Parameter(Mandatory=$True)]
         [PSCredential]$cred
         )

    Log "Trying exchange connection.."
    try{
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange `
                                 -ConnectionUri http://exch023c1cas.exchange.systemhosting.local/powershell `
                                 -Authentication Kerberos `
                                 -AllowRedirection `
                                 -Name 'Exchange' `
                                 -Credential $cred
        Import-Module (Import-PSSession -Session $Session -AllowClobber -DisableNameChecking) -Global

        }catch{
               Log "failed exchange connection.. Got error: $_"
        }
}

function Remove-Exportrequests{

    try{
        if((Get-MailboxExportRequest -Organization tsa) -notlike $null){
            Log "Removing any export requests that already exist"
            Get-MailboxExportRequest -Organization tsa | Remove-MailboxExportRequest -Confirm:$false
            
                }
            }catch{
                Log "Failed to remove export requests. Got error: $_"
         }
}

function Remove-Importrequests{

    try{
        if((Get-MailboxImportRequest -Organization tsa) -notlike $null){
            Log "Removing any Import requests that already exist"
            Get-MailboxImportRequest -Organization tsa | Remove-MailboxImportRequest -Confirm:$false
            
                }
            }catch{
                Log "Failed to remove Import requests. Got error: $_"
         }
}


function New-ContactExport{
    param(
         [Parameter(Mandatory=$True)]
         [string]$Contacts
         )



      try{
          New-MailboxExportRequest -IncludeFolders "Contacts/Systemhosting Contacts" `
                                   -FilePath "$Contacts" `
                                   -Mailbox contacts
                                   Log "Exporting contacts to pst.."
                }
    catch{
          Log "Failed to create Import requests. Got error: $_"
         }

        $Stoploop = $false
        [int]$Retrycount = "1"

        do {
	        $status   = Get-MailboxExportRequest -Organization tsa | Get-MailboxExportRequestStatistics | select -ExpandProperty status
            
		            if ($Retrycount -gt 100){
			        Write-Output "Could not export contacts."
                    Log "Could not export contacts."
			        $Stoploop = $true
		            }
		            elseif ($status -like 'Queued' -or $status -like 'InProgress'){
			                Write-Output ("Retry Count = " + "$Retrycount" + " Still trying to export contacts...")
                            Log ("Retry Count = " + "$Retrycount" + " Still trying to export contacts...")
			                Start-Sleep -Seconds 5
			                $Retrycount = $Retrycount + 1
		                    }
                            elseif($status -like 'Completed'){
                                   Write-Output ("Contacts exported successfully...")
                                   Log ("Contacts exported successfully...")
                                   $Stoploop = $true
                                  
	                        }
                }
        While ($Stoploop -eq $false)
}


function New-ContactImport{
    param(
         [Parameter(Mandatory=$True)]
         [string]$Contacts
         )

    try{
    $TSA_Mailbox = Get-Mailbox -Organization tsa | where {$_.RecipientTypeDetails -eq "UserMailbox" -and
                                                          $_.alias -notlike "hostmaster" -and
                                                          $_.alias -notlike "Vagten" -and
                                                          $_.alias -notlike "support" -and
                                                          $_.alias -notlike "test" -and
                                                          $_.alias -notlike "opsmgr" -and
                                                          $_.alias -notlike "nav2godemo" -and
                                                          $_.alias -notlike "IPVisionTSA01" -and
                                                          $_.alias -notlike "core" -and
                                                          $_.alias -notlike "connect"}
             }
  catch{
        Log "failed to get mailboxes.. Got error: $_"
       }

       if($TSA_Mailbox -notlike $null){

            foreach ($mailbox in $TSA_Mailbox){

                try{
                    New-MailboxImportRequest -Mailbox $mailbox.PrimarySmtpAddress -FilePath "$Contacts" -TargetRootFolder "/" -BatchName $mailbox.SamAccountName -Name $mailbox.alias
                    Log ("Importing the contacts to " + $mailbox.PrimarySmtpAddress + "..")
                    }
                    catch{
                          Log ("failed to create import for " + $mailbox.PrimarySmtpAddress + " Got error: $_")
                          }
	            }
        }
}




Connect-exchange -cred $cred

Remove-Exportrequests

Remove-Importrequests

New-ContactExport -Contacts $Contacts

New-ContactImport -Contacts $Contacts