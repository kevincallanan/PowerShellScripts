# ==============================================================================================
#    NAME: UpdatePowerPolicy.ps1
#    AUTHOR: Kevin Callanan
#    Released Version: 1.0
#	 Description: This script will update the PowerPolicy on servers included in .\servers.txt to High Performance		
#    WARRENTY: None - Proceed at your own risk :-) 
# ==============================================================================================

$Servers = @(Import-CSV -Path '.\servers.txt')
Start-Transcript .\transcript.log
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
	$policy = Get-CimInstance -Namespace root\cimv2\power -Class win32_PowerPlan -Filter "ElementName = 'High Performance'"
	Invoke-CimMethod -InputObject $policy -MethodName Activate
   	  }
	}

}
Stop-Transcript