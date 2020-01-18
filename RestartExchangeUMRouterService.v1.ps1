# ==============================================================================================
#    NAME: RestartExchangeUMRouterService.v1.ps1
#    AUTHOR: Kevin Callanan
#    Released Version: 1.0
#	 Description: This script will connect to each server provided issue a service restart for the
#	 Description: MSExchangeUMCR service
#    WARRENTY: None - Proceed at your own risk :-) 
# ==============================================================================================

$Servers = "Server1","Server2","Server3"
ForEach ($i in $Servers) {

    $Server = $i
    Write-Host 'Working on Server' $Server -BackgroundColor Yellow -ForegroundColor Black
    Invoke-Command -ComputerName $Server -ScriptBlock { Get-Service -Name MSExchangeUMCR | Restart-Service -Verbose}
 }