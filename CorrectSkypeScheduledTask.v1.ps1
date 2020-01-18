# ==============================================================================================
#    NAME: CorrectSkypeScheduledTask.v1.ps1
#    AUTHOR: Kevin Callanan
#    Released Version: 1.0
#	 Description: This script will connect to each server provided and correct the Working Dir for
#	 Description:  the 'PurgeHealthMontioringUserTransaction.ps1 when Skype is installed on a non-standard drive letter			
#    WARRENTY: None - Proceed at your own risk :-) 
# ==============================================================================================

$Servers = @(Import-CSV -Path '.\servers.txt')

If ($Servers.count -like 0){
Write-Host 'No Servers Listed' 
Exit
}
Else{
Write-Host "Total Server Count: " -NoNewline
Write-Host $Servers.count
}

ForEach ($i in $Servers) {

    $Server = $i.ServerName
    Write-Host 'Working on Server' $Server -BackgroundColor Yellow -ForegroundColor Black
	Invoke-Command -ComputerName $Server -ScriptBlock {
   	& { 
		$Action = New-ScheduledTaskAction -Execute "%systemroot%\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument "-ExecutionPolicy Bypass .\PurgeHealthMonitoringUserTransactions.ps1" -WorkingDirectory "E:\Program Files\Skype for Business Server 2015\Deployment\"
		Set-ScheduledTask -TaskName "\Skype for Business Server 2015\PurgeHealthMonitoringUsersTransactions" -Action $Action
   	  }
	}

}