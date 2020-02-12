$Pools = Get-CsPool |WHERE {$_.Identity -like 'skfe*'} |SELECT Identity

ForEach ($Pool in $Pools) {
    $Servers = Get-CSPool -Identity $Pool.Identity |SELECT -ExpandProperty Computers

    ForEach ($Server in $Servers){
    $WebCall = 'https://'+ $Server + '/WebTicket/WebTicketService.svc/mex'
    #Write-Host $Webcall
    $StatusDetails = Invoke-WebRequest -Uri $WebCall |SELECT StatusCode,StatusDescription,Headers
    $StatusDetails
    }
}