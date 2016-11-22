# Script to get all UPD storage for internal users on customer DFS shares..
#
# V1.0 by Jakob Strøm
#
# The script will get the UPD (single and combined) size for test and $ users on the shared platform, 
# and send the result in an email specified below..
#
# IMPORTANT! Run the script from CORP!!
# 

$cred   = Get-AutomationPSCredential -Name "adminjst"
$ExchAD = 'AD024C1EXCHGC.exchange.systemhosting.local'
$CorpAD = 'dc-03.corp.systemhosting.dk'
$mail   = @("jst@systemhosting.dk", "supporten@systemhosting.dk")

Import-Module ActiveDirectory

$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 10px; }"
$style = $style + "TD{border: 1px solid black; padding: 10px; text-align: right; }"
$style = $style + "</style>"

$logFile = "C:\Scripts\Logs\UPD_FileSizes " + (Get-Date -Format "dd-MM-yyyy_HH-mm-ss") + ".txt"

function Log([array]$text) {
	foreach ($txt in $text) {
		
		$txt | Out-File -FilePath $logFile -Encoding "UTF8" -Append
	}
}

$Customers = Get-ADOrganizationalUnit -filter * `
                                      -SearchBase "OU=Microsoft Exchange Hosted Organizations,DC=Exchange,DC=Systemhosting,DC=local" `
                                      -SearchScope onelevel `
                                      -Server $ExchAD `
                                      -Credential $cred | 
             where{$_.Name -notlike "*SVANE*" -and 
                   $_.Name -notlike "*systemhosting*" -and
                   $_.Name -notlike "*test*"} | 
             select Name, DistinguishedName

$CustomersName = $Customers | select -ExpandProperty Name
$CustomersOU   = $Customers | select -ExpandProperty DistinguishedName


$objectCollection=@()

foreach($Name in $CustomersName){

       $CustomerUsers =  Get-ADUser -Filter {Name -like "*test*" -or SamAccountName -like '*$*'} `
                                    -Properties Displayname, LastLogonDate, UserPrincipalName, SID `
                                    -SearchScope Subtree `
                                    -SearchBase "OU=$Name,OU=Microsoft Exchange Hosted Organizations,DC=Exchange,DC=Systemhosting,DC=local" `
                                    -Server $ExchAD `
                                    -Credential $cred
       $Adminusers    =  Get-ADUser -Filter {Name -like "*test*" -or SamAccountName -like '*$*'} `
                                    -Properties Displayname, LastLogonDate, UserPrincipalName, SID `
                                    -SearchScope Subtree `
                                    -SearchBase "OU=Users,OU=Admins,OU=Systemhosting,DC=Exchange,DC=Systemhosting,DC=local" `
                                    -Server $ExchAD `
                                    -Credential $cred
       $Corpusers     =  Get-ADUser -Filter {Name -like "*test*" -or SamAccountName -like '*$*'} `
                                    -Properties Displayname, LastLogonDate, UserPrincipalName, SID `
                                    -SearchScope Subtree `
                                    -SearchBase "OU=Admins,OU=SYSTEMHOSTING,dc=corp,dc=systemhosting,dc=dk" `
                                    -Server $CorpAD `
                                    -Credential $cred

            foreach($user in $CustomerUsers){
                    $UPD = Get-Item -Path ("\\customer\data\" + "$Name" + "_RDSCollection\" + "UVHD-" + $user.sid.Value + ".vhdx") -ErrorAction ignore
                    $UPDSize = $UPD.Length/1MB

                    if($UPD){
                    $object = New-Object PSObject
                    Add-Member -InputObject $object -MemberType NoteProperty -Name Customer -Value $Name
                    Add-Member -InputObject $object -MemberType NoteProperty -Name DisplayName -Value $user.DisplayName
                    Add-Member -InputObject $object -MemberType NoteProperty -Name LastLogon -Value $user.LastLogonDate
                    Add-Member -InputObject $object -MemberType NoteProperty -Name UPN -Value $user.UserPrincipalName
                    Add-Member -InputObject $object -MemberType NoteProperty -Name UPDSizeMB -Value $UPDSize
                    Add-Member -InputObject $object -MemberType NoteProperty -Name UPDPath -Value $UPD.FullName
                    $objectCollection += $object}
                    }

            foreach($user in $Adminusers){
                    $UPD = Get-Item -Path ("\\customer\data\" + "$Name" + "_RDSCollection\" + "UVHD-" + $user.sid.Value + ".vhdx") -ErrorAction ignore
                    $UPDSize = $UPD.Length/1MB

                    if($UPD){
                    $object = New-Object PSObject
                    Add-Member -InputObject $object -MemberType NoteProperty -Name Customer -Value $Name
                    Add-Member -InputObject $object -MemberType NoteProperty -Name DisplayName -Value $user.DisplayName
                    Add-Member -InputObject $object -MemberType NoteProperty -Name LastLogon -Value $user.LastLogonDate
                    Add-Member -InputObject $object -MemberType NoteProperty -Name UPN -Value $user.UserPrincipalName
                    Add-Member -InputObject $object -MemberType NoteProperty -Name UPDSizeMB -Value $UPDSize
                    Add-Member -InputObject $object -MemberType NoteProperty -Name UPDPath -Value $UPD.FullName
                    $objectCollection += $object}
                    }

            foreach($user in $Corpusers){
                    $UPD = Get-Item -Path ("\\customer\data\" + "$Name" + "_RDSCollection\" + "UVHD-" + $user.sid.Value + ".vhdx") -ErrorAction ignore
                    $UPDSize = $UPD.Length/1MB

                    if($UPD){
                    $object = New-Object PSObject
                    Add-Member -InputObject $object -MemberType NoteProperty -Name Customer -Value $Name
                    Add-Member -InputObject $object -MemberType NoteProperty -Name DisplayName -Value $user.DisplayName
                    Add-Member -InputObject $object -MemberType NoteProperty -Name LastLogon -Value $user.LastLogonDate
                    Add-Member -InputObject $object -MemberType NoteProperty -Name UPN -Value $user.UserPrincipalName
                    Add-Member -InputObject $object -MemberType NoteProperty -Name UPDSizeMB -Value $UPDSize
                    Add-Member -InputObject $object -MemberType NoteProperty -Name UPDPath -Value $UPD.FullName
                    $objectCollection += $object}
                    }
        
       } 
$objectCollection | ft 

$SUM = $objectCollection | select -ExpandProperty UPDSizeMB | Measure-Object -Sum
$TotalSum = $sum.Sum/1000
$TotalGB = "Total UPD storage used " + $TotalSum + " GB"


$report = $objectCollection | select Customer, DisplayName, LastLogon, UPN, UPDSizeMB, UPDPath | 
                              Sort-Object Customer, UPDsizeMB | ConvertTo-Html -Head $style

Send-MailMessage -SmtpServer relay.systemhosting.dk `
                 -BodyAsHtml `
                 -From driften@systemhosting.dk `
                 -To $mail `
                 -Body ($TotalGB + "$report") `
                 -Subject "UPD Storage ALL Customers"
