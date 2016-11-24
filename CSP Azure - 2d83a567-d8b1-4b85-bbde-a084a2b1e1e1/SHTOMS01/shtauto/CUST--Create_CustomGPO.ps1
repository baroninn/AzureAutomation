<#
    .DESCRIPTION
    Simple script to create a customer Default GPO. If one doesn't exist, script will create it.   

    .NOTES
        AUTHOR: Jakob Strøm
        LASTEDIT: June 13, 2016


#>



$cred = Get-AutomationPSCredential -Name "EXCHANGE--SVC_AzureAutomation"


# Import Module 
Import-Module ActiveDirectory
Import-Module GroupPolicy

$custDomain       = "customer.systemhosting.local"
$custDc           = "ad021c1custgc.customer.systemhosting.local"
$GPOName          = "_GPO_RDS_Custom"
$Customers        = Get-ADOrganizationalUnit -Filter * -SearchBase "OU=Customers,OU=SystemHosting,DC=customer,DC=systemhosting,DC=local" -SearchScope OneLevel -Server $custDc | where{$_.Name -notlike "*tsa*"}

$ErrorActionPreference = "Stop"

foreach($cust in $customers.name){
        
        $GPO = Get-GPO -Domain $custDomain -Server $custDc -All | where{$_.DisplayName -like ($cust + "_GPO_RDS_Custom*")}
        if(!$gpo){
             try{
             New-GPO -Name ($cust + "_GPO_RDS_Custom") `
                     -Comment ("Custom GPO for " + $cust) `
                     -Domain $custDomain `
                     -Server $custDc
             Write-Host "created $cust custom GPO"
                 }
                 catch{
                       $ErrorMessage = $_.Exeption.Message
                       Write-Output "Got error:`n$_"}

             try{
             New-GPLink -Name ($cust + "_GPO_RDS_Custom") `
                        -Target "OU=RDS,OU=Servers,OU=$cust,OU=Customers,OU=SystemHosting,DC=customer,DC=systemhosting,DC=local" `
                        -LinkEnabled Yes `
                        -Order 1 `
                        -Domain $custDomain `
                        -Server $custDc
                }
                catch{
                      $ErrorMessage = $_.Exeption.Message
                      Write-Output "Got error:`n$_"}

             try{
             $RDSServer = Get-ADGroup -Filter *  `
                                      -SearchBase "OU=$cust,OU=Customers,OU=SystemHosting,DC=customer,DC=systemhosting,DC=local" `
                                      -Server $custDc | 
                                      where {$_.Name -like ("$cust" + "_Organization_RDSServers")}

             $RDSUser   = Get-ADGroup -Filter *  `
                                      -SearchBase "OU=$cust,OU=Customers,OU=SystemHosting,DC=customer,DC=systemhosting,DC=local" `
                                      -Server $custDc | 
                                      where {$_.Name -like ("$cust" + "_Organization_RDSUsers")}


             Set-GPPermissions -Name ($cust + "_GPO_RDS_Custom") `
                               -TargetName $RDSServer.Name `
                               -PermissionLevel GpoApply `
                               -TargetType Group `
                               -DomainName $custDomain `
                               -Server $custDc `
                               -Replace

             Set-GPPermissions -Name ($cust + "_GPO_RDS_Custom") `
                               -TargetName $RDSUser.Name `
                               -PermissionLevel GpoApply `
                               -TargetType Group `
                               -DomainName $custDomain `
                               -Server $custDc `
                               -Replace

                        }
                        catch{
                              $ErrorMessage = $_.Exeption.Message
                              Write-Output "Got error:`n$_"}
        }
}