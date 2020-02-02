# ==============================================================================================
#    NAME: CAPI2Logging.ps1
#    AUTHOR: Kevin Callanan
#    Released Version: 1.0
#	 Description: This script will connect to each server provided and enable the CAPI2/Operational
#	 Windows Event Log located under /Applications and Sevices Logs/Microsoft/CAPI2/Operational	
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

    $Server = $i
    Write-Host 'Working on Server' $Server -BackgroundColor Yellow -ForegroundColor Black
    Invoke-Command -ComputerName $Server -ScriptBlock { 
	wevtutil sl Microsoft-Windows-CAPI2/Operational /ms:20000000
	wevtutil.exe sl Microsoft-Windows-CAPI2/Operational /e:true
    }
 }
 