# ==============================================================================================
# 
#    NAME: .\CorruptedManagedFolders.v1.0.ps1
#    AUTHOR: Kevin Callanan
#    Initial Alpha Date: 4/15/2013
#    Released Version: 1.0
#	 Relased Version Date: 4/17/2013
#    WARRENTY: None - Proceed at your own risk :-)
#    
# ==============================================================================================

param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$UserToCheck=$(throw "Please Provide Display Name of User To Check"))

#Set Job Location via script path
$JobLocation = Split-Path $MyInvocation.MyCommand.Definition -Parent -Resolve
$JobLocation = $JobLocation + "\"
$Time = Get-Date -format s
$Time = $Time -replace ":", ""

$LogUserName = Get-Mailbox $UserToCheck |SELECT DisplayName
$LogUserName = $LogUserName.DisplayName
$LogUserName = $LogUserName -replace ";", ""	
$LogUserName = $LogUserName -replace " ", ""
$LogUserName = $LogUserName -replace ",", ""
$LogUserName = $LogUserName -replace "@", ""

$LogFile = $JobLocation + "MailboxCheck_" + $LogUserName + "_" + $Time + ".log"
Start-Transcript $LogFile

#Functions	
Function Get-CorruptionEventLogs{
$Events = @()
$Events = Get-EventLog -ComputerName $ServerName -LogName Application -After $CurrentDate -Source "MSExchangeIS*" 
$Events |Sort-Object -Property "Time" -Descending  |WHERE {$_.Message -like "*$RequestIDString*"} |SELECT TimeGenerated,Source,Message,EventID
}

#Add Exchange 2010 Powershell Snapin
$null = Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue
$2007snapin = Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue
if (!$2007snapin) {
	$null = Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
	$2010snapin = Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
}

#Clear Screen
CLEAR

#Set Variables
$CurrentDate = (Get-Date).AddMinutes(-1)
$CheckUser = New-MailboxRepairRequest -Mailbox $UserToCheck -CorruptionType provisionedFolder,SearchFolder,aggregatecounts,folderview -DetectOnly

$RequestIDString = $CheckUser.RequestID.ToString()
$ServerName = $CheckUser.Server
$DatabaseName = $CheckUser.Database
$MailboxDetails = $CheckUser.Mailbox
$MessageBody = ("""*" + $RequestIDString +"*""")
Write-Host "### Running Mailbox Corruption Check ### " -ForegroundColor Red
Write-Host ""
Write-Host "RequestID: ($RequestIDString)"
Write-Host "Mailbox: ($MailboxDetails)"
Write-Host "Database: ($DatabaseName)"
Write-Host "Server: ($ServerName)"
Write-Host ""
Write-Host "Waiting 20 Seconds ..." -ForegroundColor Yellow

Start-Sleep -s 20

$EventCheck = Get-CorruptionEventLogs
$EventCheck |FL TimeGenerated,Message
$CorruptionCheck = 0

ForEach ($Event in $EventCheck){
	IF ($Event.Message -like "*Corruptions detected*"){
	Write-Host "Corruption has been detected" -ForegroundColor Black -BackgroundColor White
	Write-Host ""
	$CorruptionCheck = ($CorruptionCheck + 1)
	}
}
IF ($CorruptionCheck -gt 0 ){
$CorruptionCheckType = Read-Host "Please Type 'Yes' to Process Mailbox Repair Request..."
}

IF ($CorruptionCheckType -like "Yes"){
	$CheckUser = New-MailboxRepairRequest -Mailbox $UserToCheck -CorruptionType provisionedFolder,SearchFolder,aggregatecounts,folderview
	$RequestIDString = $CheckUser.RequestID.ToString()
	Write-Host
	Write-Host "### Fixing Corrupted Managed Folders ###" -ForegroundColor Green
	Write-Host ""
	Write-Host "RequestID: ($RequestIDString)"

	Start-Sleep -s 20

	Write-Host ""
	Write-Host "Waiting 20 Seconds ...." -ForegroundColor Yellow
	Write-Host ""
	$EventCheck = Get-CorruptionEventLogs
	$EventCheck | FL TimeGenerated,Message
	Write-Host "Corruption has been Corrected" -ForegroundColor Black -BackgroundColor White
}

IF ($CorruptionCheck -eq 0){
Write-Host "No Corruption Detected" -ForegroundColor Black -BackgroundColor White
}

Stop-Transcript