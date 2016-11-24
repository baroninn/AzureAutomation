<#
    .DESCRIPTION
    Simple script to reboot af VM from vmm-a.corp...
	The script will initiate af shutdown based on VM Name, and wait until the machine is back online.   

    .NOTES
        AUTHOR: Jakob Strøm
        LASTEDIT: June 10, 2016
#>

param ([object]$WebHookData)

$WebhookName    =   $WebhookData.WebhookName
$WebhookHeaders =   $WebhookData.RequestHeader
$WebhookBody    =   $WebhookData.RequestBody

$ServerName    = ConvertFrom-JSON -InputObject $WebHookBody

Write-Output ($ServerName + " Test write-out")

$cred = Get-AutomationPSCredential -Name "CORP--SVC_AzureAutomation"

    $scriptBlock = {
        param( 
             [String] $ServerName
             )

    $VMShutdown = Get-SCVirtualMachine $ServerName -VMMServer 'vmm-a.corp.systemhosting.dk'

    try{
        ## check VM machinestate, and shutdown if running..
        if($VMShutdown.VirtualMachineState -eq "Running"){
            Stop-SCVirtualMachine -VM $VMShutdown -Shutdown -RunAsynchronously }
        }
        catch{
              Write-Error -Message $_.Exception
              throw $_.Exception
              }


    Start-Sleep 40 -Verbose

    do{
        $VMStart = Get-SCVirtualMachine $ServerName -VMMServer 'vmm-a.corp.systemhosting.dk'
        if($VMStart.VirtualMachineState -notlike "Running"){
            Start-SCVirtualMachine -VM $VMStart -RunAsynchronously
            }
        }

    until($VMStart.Status -eq "Running")

    }
    Write-Output $ServerName
    Invoke-Command -ComputerName 'vmm-a.corp.systemhosting.dk' -ScriptBlock $scriptBlock -ArgumentList @(,$ServerName) -Credential $cred
