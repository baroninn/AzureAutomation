## Test script for creating a user on-prem
## By Jakob Strøm

param (
		[Parameter(Mandatory=$true)]
			[string] $firstname,
			
		[Parameter(Mandatory=$true)]
			[string] $lastname,
			
			[string] $City,
			
			[string] $Office,
		
       [Parameter(Mandatory=$true)]
			[string] $Password,

       [Parameter(Mandatory=$true)]
			[string] $UPN,

       [Parameter(Mandatory=$true)]
			[string] $DisplayName
	)

$logFile 		  = $env:USERPROFILE + "\Desktop\CreateUser" + (Get-Date -Format "dd-MM-yyyy_HHmmss") + ".txt"

function Log([array]$text) {
	foreach ($txt in $text) {
		
		$txt | Out-File -FilePath $logFile -Encoding "UTF8" -Append
	}
}

$AccountPassword = ConvertTo-SecureString $Password -AsPlainText -Force

try{
    New-ADUser -DisplayName $DisplayName `
               -Name $DisplayName `
               -UserPrincipalName $UPN `
               -Description 'Created with Azure' `
               -GivenName $firstname `
               -Surname $lastname `
               -Enabled $true ` `
               -Office $Office `
               -City $City `
               -AccountPassword $AccountPassword
			   Log "New user $firstname has been created"
               
			   
    }
    catch{Log $Error.exception.Message}

