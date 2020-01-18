# ==============================================================================================
#    NAME: BESFailoverStatus.v1
#    AUTHOR: Kevin Callanan
#    Released Version: 1.0
#	 Description: This script will connect to your BES SQL Database and Provide Failover Statuses for all HA pairs			
#	 Description: This script must run as an account with permissions to Read the SQL tables on the BES Database
#    WARRENTY: None - Proceed at your own risk :-) 
# ==============================================================================================

$RunTime = Get-Date

#Editable Variables
$From = "besadminonprem@contoso.com"
$To = "kevin_callanan@contoso.com"
$smtphost = "mailhub.contoso.com"
$ServerName = "sqlserver.contoso.com"
$DatabaseName = "BESMgmtOP_pdb"
$BASURL = "https://exchange-baspool.contoso.com/webconsole/login"

#SQL Query Variables 
$Query = "
SELECT UPPER([ServiceName]) AS 'ServiceName'
	,[MachineName],[SRPid],[AutoFailover]
	,[AutoFailoverExt] 
FROM [BESMgmt50].[dbo].[ServerConfig] 
WHERE ServiceName NOT LIKE MachineName"

$QueryTimeout = 120
$ConnectionTimeout = 30

$SuccessLogo = "<img id='logo-png' src='https://www.testexchangeconnectivity.com/Images/Success.png' alt='Success' />"
$SuccessWarnLogo = "<img id='logo-png' src='https://www.testexchangeconnectivity.com/Images/SuccessWarn.png' alt='SuccessWarn' />"
$ErrorLogo = "<img id='logo-png' src='https://www.testexchangeconnectivity.com/Images/Error.png' alt='Error' />"

#SQL Magic
$SQLConnection = New-Object System.Data.SqlClient.SQLConnection
$ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $ServerName,$DatabaseName,$ConnectionTimeout
$SQLConnection.ConnectionString = $ConnectionString
$SQLConnection.Open()
$QueryCMD = New-Object system.Data.SqlClient.SqlCommand($Query,$SQLConnection)
$QueryCMD.CommandTimeout = $QueryTimeout
$DataTable = New-Object system.Data.DataTable
$DataAdapter = New-Object system.Data.SqlClient.SqlDataAdapter($QueryCMD)
$DataAdapter.fill($DataTable) | Out-Null
$SQLConnection.Close()
$Results = @($DataTable)

#Build Email HTML
$EmailBody = "<html><head><meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'><title>BES HA Automated Failover Status</title></head>"
$EmailBody += "<h2>BES HA Automated Failover Status</h2>" + "<br>"
$EmailBody += "<table><center><tr><td><b>Service Name</b></td><td><b>Active Server</b></td><td><b>SRPid</b></td><td><b>AutoFailover</b></td><td><b>AutoConn Failover</b></td></tr></center>"

	ForEach ($Service in $Results)	{
		#Set Icon Type for AutoFailover Option
		IF ($Service.AutoFailover -eq 1){$AutoFailover = $SuccessLogo}
		ELSE{$AutoFailover = $ErrorLogo}
		#Set Icon Type for AutoFailoverExt Option
		IF ($Service.AutoFailoverExt -eq 1){$AutoFailoverExt = $SuccessLogo}
		ELSE{$AutoFailoverExt = $ErrorLogo}

	$EmailBody += "<tr><td>" + $Service.ServiceName + "</td><td>" + $Service.MachineName + "</td><td>" + $Service.SRPid + "</td><td><center>" + $AutoFailover + "</center></td><td><center>" + $AutoFailoverExt + "</center></td></tr>"
	}

#Complete Email HTML
$EmailBody += "</table>"
$EmailBody += "<br><a href=" + $BASURL + ">Blackberry Admin Console</a><br><br>"
$EmailBody += "<br><i>Data Generated on: " + $RunTime + "</i>"
$Subject = "BES HA Automatic Failover Status: (" + $RunTime + ")"

Send-MailMessage -To $to -Subject $Subject -Body $EmailBody  -SmtpServer $smtphost -From $from -BodyAsHtml