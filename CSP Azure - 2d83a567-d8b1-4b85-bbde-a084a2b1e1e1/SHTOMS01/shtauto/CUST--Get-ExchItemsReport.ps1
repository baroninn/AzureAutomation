<#
    .DESCRIPTION
    Creates a HTML report with information about customer mailboxes..   

    .NOTES
        AUTHOR: Jakob Strøm
        LASTEDIT: September 09, 2016


#>

$logFile  = "C:\Scripts\Logs\ItemsReport_" + (Get-Date -Format "dd-MM-yyyy_HH-mm-ss") + ".txt"

function Log([array]$text) {
	foreach ($txt in $text) {
		
		$txt | Out-File -FilePath $logFile -Encoding "UTF8" -Append
	}
}
$Mail = 'nds@systemhosting.dk'
$ErrorActionPreference = 'stop'
$cred = Get-AutomationPSCredential -Name "EXCHANGE--SVC_AzureAutomation"

#  Get Organization, only used when run manually.
#$Organisation = Read-Host "Customer initials.. (Leave blank for full report)" 

# make the connection to exchange:
Log "Trying exchange connection.."
try{
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange `
                             -ConnectionUri http://exch023c1cas.exchange.systemhosting.local/powershell `
                             -Credential $cred `
                             -Authentication Kerberos `
                             -AllowRedirection `
                             -Name 'Exchange'
    Import-Module (Import-PSSession -Session $Session -AllowClobber -DisableNameChecking) -Global

    }catch{
           Log "failed exchange connection.."
}


Import-Module ActiveDirectory

## start the script based on given parameters. Creates an object based variable which will be translated to HTML for email delivery.
$info = @()
if(!($Organisation)){
        $MBXservers = Get-MailboxDatabase
        $Customers  = Get-ADOrganizationalUnit -Filter * `
                                               -SearchBase 'OU=Microsoft Exchange Hosted Organizations,DC=exchange,DC=systemhosting,DC=local' `
                                               -SearchScope OneLevel `
                                               -Server AD025C1EXCHGC.exchange.systemhosting.local `
                                               -Credential $cred

        Log ("Found " + $customers.count + " Customers. Creating ItemsReport..")
        foreach($cust in $Customers){

            Log ("Creating report for " + $cust.name + (Get-Date -Format "dd-MM-yyyy_HH-mm-ss"))
            $Mailbox = Get-Mailbox -Organization $cust.Name | where{$_.RecipientTypeDetails -eq 'UserMailbox'}
            foreach($box in $mailbox){

                    $Folders = $box | Get-MailboxFolderStatistics
                    $BigFolders = $Folders | where{$_.ItemsInFolder -gt '30000'} | select Name, ItemsInFolder, Folderpath
                    $Biggest = $folders | Sort-Object ItemsInFolder | select -Last 1
                    $stats   = $box | Get-MailboxStatistics
                    if($BigFolders){

                    # Add each mailbox to variable with the following values:
                    $object  = New-Object PSObject
                               Add-Member -InputObject $object -MemberType NoteProperty -Name Customer -Value $cust.Name
                               Add-Member -InputObject $object -MemberType NoteProperty -Name User -Value $box.DisplayName
                               Add-Member -InputObject $object -MemberType NoteProperty -Name Initialer -Value $box.alias
                               Add-Member -InputObject $object -MemberType NoteProperty -Name Server -Value $box.ServerName
                               Add-Member -InputObject $object -MemberType NoteProperty -Name ItemsTotal -Value $stats.ItemCount
                               Add-Member -InputObject $object -MemberType NoteProperty -Name ItemsTotalSize -Value $stats.TotalItemSize
                               Add-Member -InputObject $object -MemberType NoteProperty -Name FolderCount -Value $Folders.count
                               Add-Member -InputObject $object -MemberType NoteProperty -Name FoldersOver30K -Value $BigFolders.count
                               Add-Member -InputObject $object -MemberType NoteProperty -Name BiggestFolderName -Value $Biggest.Name
                               Add-Member -InputObject $object -MemberType NoteProperty -Name BiggestFolderCount -Value $Biggest.ItemsInFolder
                               Add-Member -InputObject $object -MemberType NoteProperty -Name BiggestFolderPath -Value $Biggest.FolderPath
                               $info += $object}

                        }
            }
 }
 elseif($Organisation){
        $Mailbox = Get-Mailbox -Organization $Organisation | where{$_.RecipientTypeDetails -eq 'UserMailbox'}

            foreach($box in $Mailbox){
                    Log ("Creating report for " + $cust.name)
                    $Folders = $box | Get-MailboxFolderStatistics
                    $BigFolders = $Folders | where{$_.ItemsInFolder -gt '30000'} | select Name, ItemsInFolder, Folderpath
                    $Biggest = $folders | Sort-Object ItemsInFolder | select -Last 1
                    $stats   = $box | Get-MailboxStatistics
                    $object  = New-Object PSObject
                               Add-Member -InputObject $object -MemberType NoteProperty -Name Customer -Value $Organisation
                               Add-Member -InputObject $object -MemberType NoteProperty -Name User -Value $box.DisplayName
                               Add-Member -InputObject $object -MemberType NoteProperty -Name Initialer -Value $box.alias
                               Add-Member -InputObject $object -MemberType NoteProperty -Name Server -Value $box.ServerName
                               Add-Member -InputObject $object -MemberType NoteProperty -Name ItemsTotal -Value $stats.ItemCount
                               Add-Member -InputObject $object -MemberType NoteProperty -Name ItemsTotalSize -Value $stats.TotalItemSize
                               Add-Member -InputObject $object -MemberType NoteProperty -Name FolderCount -Value $Folders.count
                               Add-Member -InputObject $object -MemberType NoteProperty -Name FoldersOver30K -Value $BigFolders.count
                               Add-Member -InputObject $object -MemberType NoteProperty -Name BiggestFolderName -Value $Biggest.Name
                               Add-Member -InputObject $object -MemberType NoteProperty -Name BiggestFolderCount -Value $Biggest.ItemsInFolder
                               Add-Member -InputObject $object -MemberType NoteProperty -Name BiggestFolderPath -Value $Biggest.FolderPath
                               $info += $object

                        }
            }


## Generate HTML Table styles
$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #58FA82; padding: 10px; }"
$style = $style + "TD{border: 1px solid black; padding: 10px; text-align: right; }"
$style = $style + "</style>"

Log "Trying to create HTML report and sent it by mail.."
$Report = $info | Sort-Object Server, BiggestFolderCount | ConvertTo-Html -Head $style

## If report has been generated, send it to specified mail address.
if($info){

        Send-MailMessage -SmtpServer "relay.systemhosting.dk" `
                         -BodyAsHtml `
                         -From "jst@systemhosting.dk" `
                         -To "$Mail" `
                         -Body "$Report" `
                         -Subject "Exchange user overview"
}