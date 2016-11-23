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
                             -CommandName Get-Mailbox, Enable-Mailbox, Set-Mailbox, Get-Group) -Global

    }catch{
           Write-Output "failed exchange connection.."
}

$members = Get-Group ICE_O365_Archive_1Year | select members

if($session){

    foreach($member in $members.Members){
        $mailbox = Get-Mailbox "$member" -Filter {ArchiveStatus -eq "None" -and RecipientTypeDetails -eq "UserMailbox"}

        foreach($i in $mailbox){

                Enable-Mailbox -Identity $i.Identity -Archive -ArchiveName ($i.alias + " - Archive")
                Set-Mailbox -Identity $i.Identity -RetentionPolicy "1 year archive"
                Write-output "Archive and retention enabled for $i"
                }
        }
}

Get-PSSession | Remove-PSSession

