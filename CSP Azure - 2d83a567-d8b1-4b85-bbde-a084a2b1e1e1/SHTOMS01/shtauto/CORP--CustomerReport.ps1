 ###############################################
##
## Customer Report script.
##
## V1.0 By Jakob Strøm
## V1.1 added logic for disabled users not counting in report..
## V1.2 Added SCVMM connection and details about customers Servers..
## V1.3 Added logic for test users. Every username with "*TEST*" in its name will not be counted.
## V1.4 Rewritten parts of the script to support Azure Automation.
## 

# Import modules and global variables
$ErrorActionPreference = 'stop'
$VMMHost               = 'vmm-a.corp.systemhosting.dk'
#Get-AutomationPSCredential -Name "CORP--SVC_AzureAutomation"

Import-Module ActiveDirectory -Global
Import-Module GroupPolicy -Global
Import-Module virtualmachinemanager -Cmdlet "Get-SCCloud", "Get-VM" -Global -ErrorAction SilentlyContinue

$customers=@()
$ICE = New-Object PSObject
       Add-Member -InputObject $ICE -MemberType NoteProperty -Name Customer -Value 'ICE'
       Add-Member -InputObject $ICE -MemberType NoteProperty -Name Domain -Value 'corp.icepower.dk'
$BOH = New-Object PSObject
       Add-Member -InputObject $BOH -MemberType NoteProperty -Name Customer -Value 'BOH'
       Add-Member -InputObject $BOH -MemberType NoteProperty -Name Domain -Value 'corp.buch-holm.dk'
$PRV = New-Object PSObject
       Add-Member -InputObject $PRV -MemberType NoteProperty -Name Customer -Value 'PRV'
       Add-Member -InputObject $PRV -MemberType NoteProperty -Name Domain -Value 'corp.provinord.dk'
$SGC = New-Object PSObject
       Add-Member -InputObject $SGC -MemberType NoteProperty -Name Customer -Value 'SGC'
       Add-Member -InputObject $SGC -MemberType NoteProperty -Name Domain -Value 'corp.thescandinavian.dk'
$ASG = New-Object PSObject
       Add-Member -InputObject $ASG -MemberType NoteProperty -Name Customer -Value 'ASG'
       Add-Member -InputObject $ASG -MemberType NoteProperty -Name Domain -Value 'asgdom.local'

$customers += $ICE, $BOH, $PRV, $SGC, $ASG
# Test single customer
#$customers += $ASG

foreach($i in $customers){

    ## Customer specific variables
    $Customer     = $i.Customer
    $Domain       = $i.Domain
    if($Customer -eq 'ASG'){$Domdc = (($Customer) + "-dc01." + ($Domain))}else{
    $Domdc        = (($Customer) + "-dc-01." + ($Domain))}
    $Account      = ("SVC_" + "$Customer" + "_CustomerReport@" + "$Domain")
    $SVC_Password = 'ijhu45%¤/2D42!"v87¤g3¤%&rd¤'
    $SVC_secPass  = ConvertTo-SecureString -String $SVC_Password -AsPlainText -Force
    $cred         = new-object -typename System.Management.Automation.PSCredential -argumentlist $Account, $SVC_secPass
    $logFile      = "C:\Scripts\Logs\CustomerReport_" + "$Customer _" + (Get-Date -Format "dd-MM-yyyy_HH-mm-ss") + ".txt"

    function Log([array]$text) {
	    foreach ($txt in $text) {
		
		    $txt | Out-File -FilePath $logFile -Encoding "UTF8" -Append
	    }
    }
    Write-Output ("Getting ADdomain for customer " + $customer)
    try{
        $domdn  = Get-ADDomain -Server $Domdc -Credential $cred | select -ExpandProperty DistinguishedName
        $domUPN = Get-ADDomain -Server $Domdc -Credential $cred | select -ExpandProperty Forest

        $Cloud  = Get-SCCloud -VMMServer $VMMHost | where{$_.Name -like "$Customer*"}
        $VMs    = Get-VM -VMMServer $VMMHost -Cloud $Cloud | where {$_.Status -eq "Running" -and $_.Name -notlike "*mpgw*"}
        }catch{
            Write-Output ($customer + " Error getting backend Info (AD, SCVMM, etc..")
            Log ($customer + " Error getting backend Info (AD, SCVMM, etc..")
        }

    ## Mail settings
    #$SMTP     = (($customer) + "-exch-01." + ($domUPN)) <- Deleted as we are sending from Corp (relay.systemhosting.dk)
    $SMTPFrom = ("SHReport@" + $domUPN)
    $SMTPTo   = "ahf@systemhosting.dk" ## Emails can be seperated by ,  "@","@","@"

    ## Generate HTML Table styles
    $style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
    $style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
    $style = $style + "TH{border: 1px solid black; background: #58FA82; padding: 10px; }"
    $style = $style + "TD{border: 1px solid black; padding: 10px; text-align: right; }"
    $style = $style + "</style>"

    ## Import remote Exchange session
    try{
        if($Customer -eq 'ASG'){$Mailsession = New-PSSession -Name Microsoft.Exchange -ConnectionUri ("http://" + $customer + "-EXCH01." + $domUPN + "/powershell") -ConfigurationName Microsoft.Exchange -Credential $cred}else{
        $Mailsession = New-PSSession -Name Microsoft.Exchange -ConnectionUri ("http://" + $customer + "-EXCH-01." + $domUPN + "/powershell") -ConfigurationName Microsoft.Exchange -Credential $cred}
        Import-PSSession $Mailsession -AllowClobber -CommandName Get-Mailbox, Get-MailboxStatistics
        }catch{
        Write-Output ($customer + " Mail server not accessible")
        Log ($customer + " Mail server not accessible")
    
       }

    ## Get Full/Light users..
    try{
        $FullMembers   = Get-ADGroupMember -Identity "G_FullUsers" -Server $Domdc -Credential $cred | Sort Name | Select Name, SamAccountName
        $fullusers = @()
        foreach($user in $FullMembers){
        
               $u = Get-ADUser -Properties Enabled,DisplayName $user.SamAccountName -Server $Domdc -Credential $cred | where{$_.DisplayName -notlike "*test*"}
               if($u.enabled -eq $true){
                    $FullUsers += $u
                    }
               }
        $LightMembers  = Get-ADGroupMember -Identity "G_LightUsers" -Server $Domdc -Credential $cred | Sort name | select name, SamAccountName
        $Lightusers = @()
        foreach($user in $LightMembers){
        
               $u = Get-ADUser -Properties Enabled,DisplayName $user.SamAccountName -Server $Domdc -Credential $cred | where{$_.DisplayName -notlike "*test*"}
               if($u.enabled -eq $true){
                    $LightUsers += $u
                    }
               }
    }catch{
          Write-Output "Got error:`n$_"
          Log ($customer + " Error getting backend Info (AD, SCVMM, etc.. Got error:`n$_")
          }

    $Servers     = Get-ADComputer -Filter * -Server $Domdc -Properties OperatingSystem -Credential $cred | 
                   where{$_.Name -like "$Customer-*" -and 
                         $_.name -notlike "*MPGW*" -and 
                         $_.OperatingSystem -like "Windows Server*"} | 
                   sort name | select name, DNSHostName, OperatingSystem

    ## Create collection with Full user information
    $objectCollectionFULL=@()
        foreach($User in $FullUsers){

            try{
                $mailboxFull = ''
                $mailboxFull = Get-Mailbox -Identity $user.SamAccountName -ErrorAction SilentlyContinue
                $mailboxFULLStats = Get-Mailbox -Identity $user.SamAccountName | Get-MailboxStatistics | Select-Object -Property DisplayName,@{label="Mailbox MB";expression={[int64](([int64]($_.TotalItemSize -split '[\( ]')[3])/1048576)}} -ErrorAction SilentlyContinue
                }
                catch{
                Write-Output ($customer + " user " + $user.name + " does not appear to have a mailbox")
                Log ($customer + " user " + $user.name + " does not appear to have a mailbox")}

            if($mailboxFull.PrimarySmtpAddress -ne $null){
            $object = New-Object PSObject
                        Add-Member -InputObject $object -MemberType NoteProperty -Name User -Value $user.name
                        Add-Member -InputObject $object -MemberType NoteProperty -Name Initialer -Value $user.SamAccountName
                        Add-Member -InputObject $object -MemberType NoteProperty -Name Email -Value $mailboxFull.PrimarySmtpAddress
                        Add-Member -InputObject $object -MemberType NoteProperty -Name MailboxSize -Value $mailboxFullStats.'Mailbox MB'
                        Add-Member -InputObject $object -MemberType NoteProperty -Name MailboxLimit -Value $mailboxFull.ProhibitSendReceiveQuota
                        $objectCollectionFULL += $object}
                   else{
                        $object = New-Object PSObject
                                  Add-Member -InputObject $object -MemberType NoteProperty -Name User -Value $user.name
                                  Add-Member -InputObject $object -MemberType NoteProperty -Name Initialer -Value $user.SamAccountName
                                  Add-Member -InputObject $object -MemberType NoteProperty -Name MailboxSize -Value "0"
                                  Add-Member -InputObject $object -MemberType NoteProperty -Name MailboxLimit -Value "NoMailbox"
                                  $objectCollectionFULL += $object}
            }

    ## Create collection with "Light user/Mail only" information
    $objectCollectionLIGHT=@()
        foreach($User in $LightUsers){

            try{
                $mailboxLIGHT = ''
                $mailboxLIGHT = Get-Mailbox -Identity $user.SamAccountName
                $mailboxLIGHTStats = Get-Mailbox -Identity $user.SamAccountName | Get-MailboxStatistics | Select-Object -Property DisplayName,@{label="Mailbox MB";expression={[int64](([int64]($_.TotalItemSize -split '[\( ]')[3])/1048576)}}
                }
                catch{Write-Output ($Customer + " user " + $user.name + " does not appear to have a mailbox")
                      Log ($Customer + " user " + $user.name + " does not appear to have a mailbox")}
            

            if($mailboxLIGHT.PrimarySmtpAddress -ne $null){
            $object = New-Object PSObject
                        Add-Member -InputObject $object -MemberType NoteProperty -Name User -Value $user.name
                        Add-Member -InputObject $object -MemberType NoteProperty -Name Initialer -Value $user.SamAccountName
                        Add-Member -InputObject $object -MemberType NoteProperty -Name Email -Value $mailboxLIGHT.PrimarySmtpAddress
                        Add-Member -InputObject $object -MemberType NoteProperty -Name MailboxSize $mailboxLIGHTStats.'Mailbox MB'
                        Add-Member -InputObject $object -MemberType NoteProperty -Name MailboxLimit -Value $mailboxLIGHT.ProhibitSendReceiveQuota
                        $objectCollectionLIGHT += $object}
                    else{
                         $object = New-Object PSObject
                        Add-Member -InputObject $object -MemberType NoteProperty -Name User -Value $user.name
                        Add-Member -InputObject $object -MemberType NoteProperty -Name Initialer -Value $user.SamAccountName
                        Add-Member -InputObject $object -MemberType NoteProperty -Name MailboxSize -Value "0"
                        Add-Member -InputObject $object -MemberType NoteProperty -Name MailboxLimit -Value "NoMailbox"
                        $objectCollectionLIGHT += $object}
                    
            }

    ### Measure the total Mailbox storage in GB
    if($objectCollectionLIGHT.MailboxSize -ne $null){
            $LightSUM = $objectCollectionLIGHT | select -ExpandProperty MailboxSize | Measure-Object -Sum
            }
            else{
                Write-Output "No Light mailboxes"
                Log "No Light mailboxes"
                }


    if($objectCollectionFULL.MailboxSize -ne $null){
            $FullSUM = $objectCollectionFULL | select -ExpandProperty MailboxSize | Measure-Object -Sum
            }
            else{
                Write-Output "No Full mailboxes"
                Log "No Full mailboxes"
                }

    if($LightSUM -or $FullSUM){
        
            $sum = $LightSUM.sum + $FullSUM.sum
            $TotalSum = $sum/1000
            $TotalMailStorage = "Total Mailbox storage used " + $TotalSum + " GB"
            }
            else{
                Write-Output "Sum not calculated when no mailboxes"
                Log "Sum not calculated when no mailboxes"
                }


    ## Create collection with server information
    $objectCollectionVMs=@()
        foreach($VM in $VMs){

                $object = New-Object PSObject
                            Add-Member -InputObject $object -MemberType NoteProperty -Name 'Server Name' -Value $vm.Name
                            if($vm.DynamicMemoryMaximumMB -ne $null){
                            Add-Member -InputObject $object -MemberType NoteProperty -Name 'Ram GB' -Value ([math]::truncate($vm.DynamicMemoryMaximumMB / 1024))
                            }else{
                            Add-Member -InputObject $object -MemberType NoteProperty -Name 'Ram GB' -Value ([math]::truncate($vm.MemoryAssignedMB / 1024))
                            }
                            Add-Member -InputObject $object -MemberType NoteProperty -Name 'CPU' -Value $vm.CPUCount
                            Add-Member -InputObject $object -MemberType NoteProperty -Name 'OS' -Value $vm.OperatingSystem
                            $objectCollectionVMs += $object
                }

    ## All Storage stats
    $objectCollectionStorage=@()
        foreach($server in $servers){
            try{
                    $session = New-CimSession -ComputerName ($server.name + "." +  "$domUPN") -Name $server.name -Authentication Negotiate -Credential $cred

                    $volumes = Get-CimInstance -CimSession $session -Query "SELECT * FROM Win32_Volume" | 
                               where{$_.Capacity -notlike $null -and 
                                     $_.Name -notlike "C:\" -and 
                                     $_.Label -notlike "User Disk" -and 
                                     $_.Label -notlike "SwapDisk"} | 
                               select Name, Label, PSComputername, @{label="Capacity";expression={[int64](([int64]($_.Capacity)/1073741824))}}

                    foreach($volume in $volumes){
          
                              $object = New-Object PSObject
                                        Add-Member -InputObject $object -MemberType NoteProperty -Name Server -Value $server.name
                                        Add-Member -InputObject $object -MemberType NoteProperty -Name 'DiskSize GB' -Value $volume.Capacity
                                        Add-Member -InputObject $object -MemberType NoteProperty -Name Disk -Value $volume.Name
                                        Add-Member -InputObject $object -MemberType NoteProperty -Name DiskName -Value $volume.Label
                                        $objectCollectionStorage += $object
                                                }
                                Get-CimSession -Name $server.name | Remove-CimSession
                            
                                        }catch{
                                               Write-Output ("Server " +  $server.name + " not accessible")
                                               Log ("Server " +  $server.name + " not accessible")}
                                              }

    $objectCollectionStorageTotalSize = $objectCollectionStorage.'DiskSize GB' | Measure-Object -Sum

    ## Create collection with overall statestics information
    $objectCollectionStats=@()
          
              $object = New-Object PSObject
                        Add-Member -InputObject $object -MemberType NoteProperty -Name 'Full Users' -Value $FullUsers.count
                        Add-Member -InputObject $object -MemberType NoteProperty -Name 'Light Users' -Value $LightUsers.count
                        if($FullSUM -ne $null){
                            Add-Member -InputObject $object -MemberType NoteProperty -Name 'Total Mailbox Size GB' -Value $TotalSum
                                            }
                                            else{
                                                Add-Member -InputObject $object -MemberType NoteProperty -Name 'Total Mailbox Size GB' -Value "No mailboxes"
                                                }
                        Add-Member -InputObject $object -MemberType NoteProperty -Name 'Server Count' -Value $Servers.Count
                        Add-Member -InputObject $object -MemberType NoteProperty -Name 'Total Storage GB' -Value $objectCollectionStorageTotalSize.Sum
                        $objectCollectionStats += $object


    ## Convert collections to HTML reports
    $FullUsersreport = $objectCollectionFULL | select User, Initialer, Email, MailboxSize, MailboxLimit | 
                                               Sort-Object User | ConvertTo-Html -Head $style

    $LightUsersreport = $objectCollectionLIGHT | select User, Initialer, Email, MailboxSize, MailboxLimit | 
                                                 Sort-Object User | ConvertTo-Html -Head $style

    if($objectCollectionFULL.MailboxSize -eq $Null){
             $FullUsersreport = $objectCollectionFULL | select User, Initialer | Sort-Object User | ConvertTo-Html -Head $style
             }
             else{
                  $FullUsersreport = $objectCollectionFULL | select User, Initialer, Email, MailboxSize, MailboxLimit | 
                                                 Sort-Object User | ConvertTo-Html -Head $style
                 }

    if($objectCollectionLIGHT.MailboxSize -eq $Null){
             $LightUsersreport = $objectCollectionLIGHT | select User, Initialer | Sort-Object User | ConvertTo-Html -Head $style
             }
             else{
                  $LightUsersreport = $objectCollectionLIGHT | select User, Initialer, Email, MailboxSize, MailboxLimit | 
                                                 Sort-Object User | ConvertTo-Html -Head $style
                 }

                    

    $StatsReport    = $objectCollectionStats | select 'Full Users', 'Light Users', 'Total Mailbox Size GB', 'Server Count', 'Total Storage GB' | ConvertTo-Html -Head $style
    $storageReport  = $objectCollectionStorage | select Server, Disk, 'DiskSize GB', DiskName | ConvertTo-Html -Head $style
    $serversReport  = $objectCollectionVMs | Select 'Server Name', CPU, 'Ram GB', OS | Sort-Object 'Server Name' | ConvertTo-Html -Head $style

    ## Send HTML mail with report
    try{
    Send-MailMessage -SmtpServer "relay.systemhosting.dk" `
                     -BodyAsHtml `
                     -From "SHReport@systemhosting.dk" `
                     -To $SMTPTo `
                     -Cc "jst@systemhosting.dk" `
                     -Body (("$StatsReport")  + 
                           ("FULL Users : " + $FullUsers.count + $FullUsersreport) +  
                           ("LIGHT Users :" + $LightUsers.count + $LightUsersreport) + 
                           ("Servers: ") + 
                           ($serversReport) + 
                           ("Server Storage") + 
                           ($storageReport)) `
                     -Subject (($Customer.ToUpper()) + " Customer Report")
                     Log "Trying to send mail report"
        }catch{
              $ErrorMessage = $_.Exeption.Message
              Write-Output "Got error:`n$_"
              Log "Got error:`n$_"
              }
}