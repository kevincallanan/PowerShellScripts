$Pools = Get-CsPool | WHERE {$_.Services -like "Registrar:*"}|Sort-Object Identity | SELECT Identity

ForEach ($Pool in $Pools) {
    $Servers = Get-CSPool -Identity $Pool.Identity |SELECT -ExpandProperty Computers
    Write-Host $Pool.Identity -ForegroundColor Yellow -BackgroundColor Black
    ForEach ($Server in $Servers){
    $WebCall = 'https://'+ $Server + '/WebTicket/WebTicketService.svc/mex'
    #Write-Host $Webcall
    $StatusDetails = Invoke-WebRequest -Uri $WebCall |SELECT StatusCode,StatusDescription,Headers
    $StatusDetails
    }
}