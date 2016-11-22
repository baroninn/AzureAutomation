$cred = Get-AutomationPSCredential -Name "IceAdminO365"

try{
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange `
                             -ConnectionUri "https://outlook.office365.com/powershell-liveid/" `
                             -Credential $cred `
                             -Authentication Basic `
                             -AllowRedirection `
                             -Name 'Exchange'
    Import-Module (Import-PSSession -Session $Session `
                                    -AllowClobber `
                                    -DisableNameChecking `
                                    -CommandName Get-Mailbox, Enable-Mailbox, Set-Mailbox) -Global

    }catch{
           Write-Output "failed exchange connection.."
}

if($session){

    $mailbox = Get-Mailbox -Filter {ArchiveStatus -eq "None" -and RecipientTypeDetails -eq "UserMailbox"}

    foreach($i in $mailbox){

            Enable-Mailbox -Identity $i.Identity -Archive -ArchiveName ($i.alias + " - Archive")
            Set-Mailbox -Identity $i.Identity -RetentionPolicy "2 year archive"
            Write-output "Archive and retention enabled for $i"
            
            }
}

Get-PSSession | Remove-PSSession