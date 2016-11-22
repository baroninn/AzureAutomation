<#
    .DESCRIPTION
    Simple script to reboot af VM from vmm-a.corp...
	The script will initiate af shutdown based on VM Name, and wait until the machine is back online.   

    .NOTES
        AUTHOR: Jakob Strøm
        LASTEDIT: June 10, 2016
#>

param (
       	[object]$WebhookData,
		[string] $VMName
      )

$cred = Get-AutomationPSCredential -Name 'adminjst'


## Scriptblock to send to VMM server..

$scriptBlock = {
                    param( 
                    [String] $VMName 
                         )
            $VMShutdown = Get-SCVirtualMachine $VMName -VMMServer 'vmm-a.corp.systemhosting.dk'
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
                $VMStart = Get-SCVirtualMachine $VMName -VMMServer 'vmm-a.corp.systemhosting.dk'
                if($VMStart.VirtualMachineState -notlike "Running"){
                    Start-SCVirtualMachine -VM $VMStart -RunAsynchronously
                    }
                }
            until($VMStart.Status -eq "Running")
            }

# If runbook was called from Webhook, WebhookData will not be null.
    if ($WebhookData -ne $null) {   

        # Collect properties of WebhookData
        $WebhookName    =   $WebhookData.WebhookName
        $WebhookHeaders =   $WebhookData.RequestHeader
        $WebhookBody    =   $WebhookData.RequestBody

        # Collect individual headers. VMList converted from JSON.
        $From = $WebhookHeaders.From
        $Servernames = ConvertFrom-Json -InputObject $WebhookBody
        Write-Output "Runbook started from webhook $WebhookName by $From."


        # Start each virtual machine
        foreach ($VM in $Servernames)
        {
            $VMName = $VM.Name
            Invoke-Command -ComputerName 'vmm-a.corp.systemhosting.dk' -ScriptBlock $scriptBlock -ArgumentList @(,$VMName) -Credential $cred
        }
} else {
            Invoke-Command -ComputerName 'vmm-a.corp.systemhosting.dk' -ScriptBlock $scriptBlock -ArgumentList @(,$VMName) -Credential $cred
        }

