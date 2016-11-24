<#
    .DESCRIPTION
    Simple script to set modify rights on all GPO's in the customer domain.
    The Script will not set the permission on base/template GPO's..
        

    .NOTES
        AUTHOR: Jakob Strøm
        LASTEDIT: June 10, 2016


#>

$cred = Get-AutomationPSCredential -Name "EXCHANGE--SVC_AzureAutomation"
# Import Modules
Import-Module ActiveDirectory
Import-Module GroupPolicy

$custDomain       = "customer.systemhosting.local"
$custDc           = "ad021c1custgc.customer.systemhosting.local"
$AccessLevel25    = "Access_Level_25"
$AccessLevel30    = "Access_Level_30"
	


$GPOS = Get-GPO -Server $custDc -Domain $custDomain -All | 
        where {$_.Displayname -like "*GPO*" -and
               $_.Displayname -notlike "*Base_GPO*" -and
               $_.Displayname -notlike "Template_GPO*" -and 
               $_.Displayname -notlike "TSA_GPO*"} | Sort-Object DisplayName

$Level25 = $(Get-ADGroup $AccessLevel25 -Server $custDc).samaccountname
$Level30 = $(Get-ADGroup $AccessLevel30 -Server $custDc).samaccountname

foreach ($GPO in $GPOS)

        {
        $GPOName = $GPO.DisplayName
        try{
           Set-GPPermissions -DomainName $custDomain -Guid $GPO.Id -TargetName $Level25 -TargetType Group -PermissionLevel GpoEdit -Server $custDc -Replace -errorvariable $SET_GP
           Set-GPPermissions -DomainName $custDomain -Guid $GPO.Id -TargetName $Level30 -TargetType Group -PermissionLevel GpoEditDeleteModifySecurity -Server $custDc -Replace -errorvariable $SET_GP
		   }
           catch{
                  Write-Host "$SET_GP"
                }
        }