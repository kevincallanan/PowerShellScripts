# ==============================================================================================
# 
#    NAME: .\EWSUsageChecker.v2.0.ps1
#    AUTHOR: Kevin Callanan
#    Initial Alpha Date: 5/24/2013
#    Released Version: 2.0
#	 Released Date: 5/27/2013
#	 Description: This script will pull the logs from all Exchange CAS servers and filter down any EWS usage outside of approved applications.
#    WARRANTY: None - Proceed at your own risk :-)
#    
# ==============================================================================================

CLEAR 
Write-Host "Running EWSUsageChecker"

#Add Exchange 2010 Powershell Snapin
$null = Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue
$2007snapin = Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue
if (!$2007snapin) {
	$null = Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
	$2010snapin = Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
}

#Editable Variables
$From = "from@contoso.com"
$To = "kevin@contoso.com"
$SMTPhost = "smtp.contoso.com"
$Subject = "Unauthorized EWS Application Usage" 
$IISLogPath = "\C$\inetpub\logs\LogFiles\W3SVC1\" #IIS Log Share Path


#Set Job Location via script path
$JobLocation = Split-Path $MyInvocation.MyCommand.Definition -Parent -Resolve
$JobLocation = $JobLocation + "\"

#Start Transcript
$Time = Get-Date -format s
$Time = $Time -replace ":", ""
$LogFile = $JobLocation + "EWSUsageChecker_" + $Time + ".log"
Start-Transcript $LogFile

#Get Client Access Servers 
$Servers = Get-ClientAccessServer 

#Build IIS Log Path
$IISLogTime = Get-Date
$IISLogTime = $IISLogTime.AddDays(-1)
$LogDateTime = Get-Date $IISLogTime -format "yyMMdd"

#Export Log Data Path
$ExportLog = $JobLocation + "Export_u_ex" + $LogDateTime + "_" + $Time + ".log"

#This will process each log file for all CAS servers and export a limited set of data based on the IF statement below.
ForEach ($Server in $Servers){
	$Count = 0
	$ServerLog = "\\" + $Server + $IISLogPath + "u_ex" + $LogDateTime + ".log"
	Write-Host "Log: $ServerLog"
	$Reader = [System.IO.File]::OpenText($ServerLog)
	try {
		for(;;) {
			$line = $Reader.ReadLine()
			IF ($line -eq $null) { break }
			#Exclude data from export log
			IF (($line -like "*EWS/*") -and`
				($line -notlike "*Microsoft*") -and`
				($line -notlike "*Mozilla/*") -and`
				($line -notlike "*OwaProxy*") -and`
				($line -notlike "*MSRPC*") -and`
				($line -notlike "*\extest*") -and`
				($line -notlike "*ASProxy/CrossSite*") -and`
				($line -notlike "*- 80 -*") -and`
				($line -notlike "*- 443 -*") -and`
				($line -notlike "*fe80::*")){
				
				$line | Out-File -FilePath $ExportLog -Append
				$count =+ $count + 1
			}
		}
	}
	finally {
		$Reader.Close()
		Write-Host "Output Items: $Count"
	}
}

#Build E-mail Body Header
$EMSUsageEmailBody = "<html><head><meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'><title>$Subject</title></head>"
$EMSUsageEmailBody += "<h2>$Subject</h2>"

#Import Extracted Data
$IISLogImport += Import-CSV -delimiter " " -path $ExportLog -Header date,time,s-ip,cs-method,cs-uri-stem,cs-uri-query,s-port,cs-username,c-ip,csUser-Agent,sc-status,sc-substatus,sc-win32-status,time-taken
	
#Select Unique User-Agent From Exported Log Data
$UniqueUserAgents = @($IISLogImport |`
	WHERE {$_."cs-uri-stem" -like "/EWS/*"}|`
	WHERE {$_."cs-UserName" -like "*\*"} |`
	WHERE {$_."csUser-Agent" -notlike "*Microsoft*"}|`
	WHERE {$_."csUser-Agent" -notlike "Mozilla*"}|`
	WHERE {$_."csUser-Agent" -notlike "ExchangeServicesClient*"}|`
	WHERE {$_."csUser-Agent" -notlike "OwaProxy*"} |`
	WHERE {$_."csUser-Agent" -notlike "-"}|`
	WHERE {$_."csUser-Agent" -notlike "MSRPC"}|`
	WHERE {$_."cs-UserName" -notlike "*\extest*"}|`
	SELECT csUser-Agent -unique | Sort-Object "csUser-Agent" -Descending)

#Output Offending Users Based on User-Agent
ForEach ($UserAgent in $UniqueUserAgents){
	$UserAgentName = $UserAgent."csUser-Agent"
	
	
	#Add E-mail Table
	$EMSUsageEmailBody += "<br><b>User-Agent: $UserAgentName</b>"
    $EMSUsageEmailBody += "<br>----------------------------------------<br>"
	$EMSUsageEmailBody += "<table><tr><td><b>Display Name</b></td><td><b>E-mail Address</b></td></tr>"
	Write-Host "EWS User-Agent: $UserAgentName"
	
	#Generate a List of Users based on User-Agent
	$ActiveUsers = @($IISLogImport |`
		WHERE {$_."cs-uri-stem" -like "/EWS/*"}|`
		WHERE {$_."cs-UserName" -like "*\*"} |`
		WHERE {$_."csUser-Agent" -notlike "*Microsoft*"}|`
		WHERE {$_."csUser-Agent" -notlike "Mozilla*"}|`
		WHERE {$_."csUser-Agent" -notlike "ExchangeServicesClient*"}|`
		WHERE {$_."csUser-Agent" -notlike "OwaProxy*"} |`
		WHERE {$_."csUser-Agent" -notlike "-"}|`
		WHERE {$_."csUser-Agent" -notlike "MSRPC"}|`
		WHERE {$_."cs-username" -notlike "*\extest*"}|`
		WHERE {$_."csUser-Agent" -like $UserAgentName}|`
		SELECT cs-username -unique)

	#Look up users Display Name & Primary E-mail Address
	ForEach ($User in $ActiveUsers){
		$UserAccount = $User."cs-username"
		$Details = Get-Mailbox -Identity $UserAccount | SELECT DisplayName,PrimarySmtpAddress
		$EMSUsageEmailBody += "<tr><td>" + $Details.DisplayName + "</td><td>" + $Details.PrimarySmtpAddress + "</td></tr>"
		Write-Host $Details.DisplayName
	}
	#Complete E-mail Table
	$EMSUsageEmailBody += "</table><br>"
}

#Complete E-mail Body
$IISLogRunTime = ( Get-Date $IISLogTime -uformat "%m-%d-%Y" ).ToString()
$EMSUsageEmailBody += "<br><i>Data Generated based IIS Logs from: " + $IISLogRunTime + "</i>"

#Send E-mail
Write-Host "Sending E-mail ... "
Send-MailMessage -To $to -Subject $Subject -Body  $EMSUsageEmailBody  -SmtpServer $SMTPhost -From $from -BodyAsHtml

Stop-Transcript 