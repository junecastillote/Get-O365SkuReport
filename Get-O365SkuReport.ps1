<#	
	.NOTES
	===========================================================================
	 Created on:   	20-December-2018
	 Created by:   	June Castillote
					june.castillote@gmail.com
	 Filename:     	Get-O365SkuReport.ps1
	 Version:		1.0 (20-December-2018)
	===========================================================================

	.DESCRIPTION
		For more details and usage instruction, please visit the link:
		https://www.lazyexchangeadmin.com/2019/01/Get-O365SkuReport.html
		https://github.com/junecastillote/Get-O365SkuReport		
		
		
	.EXAMPLE
		.\Get-O365SkuReport.ps1

#>
$scriptVersion = "1.0"
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
[xml]$config = Get-Content "$($script_root)\config.xml"
#debug-----------------------------------------------------------------------------------------------
[boolean]$enableDebug = $config.options.enableDebug
#----------------------------------------------------------------------------------------------------
#Mail------------------------------------------------------------------------------------------------
[boolean]$sendReport = $config.options.SendReport
[string]$tenantName = ""
[string]$fromAddress = $config.options.fromAddress
[string]$toAddress = $config.options.toAddress
[string]$smtpServer = "smtp.office365.com"
[int]$smtpPort = "587"
[string]$mailSubject = "Office 365 License Availability Report"
#----------------------------------------------------------------------------------------------------
#Housekeeping----------------------------------------------------------------------------------------
$enableHousekeeping = $true #Set this to false if you do not want old backups to be deleted
$daysToKeep = 60
#----------------------------------------------------------------------------------------------------
$Today=Get-Date
[string]$filePrefix = '{0:dd-MMM-yyyy_hh-mm_tt}' -f $Today
$logPath = "$($script_root)\Logs"
$logFile = "$($logPath)\DebugLog_$($filePrefix).txt"
$reportPath = "$($script_root)\Reports"
$reportFile="$($reportPath)\skuMon_$($filePrefix).csv"
$htmlFile="$($reportPath)\skuMon_$($filePrefix).html"
$skus = import-Csv "$($script_root)\Account_Sku.csv"

#Create folders if not found
if (!(Test-Path $logPath)) {New-Item -ItemType Directory -Path $logPath}
if (!(Test-Path $reportPath)) {New-Item -ItemType Directory -Path $reportPath}

Function New-MsOLSession {
    [CmdletBinding()]
    param(
        [parameter(mandatory=$true)]
        [PSCredential] $msolCredential
    )

	if (Get-Module msOnline -ListAvailable)	{
		Import-Module MSOnline
		Connect-MsolService -Credential $msolCredential
	}
	else {
		Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": MSonline Module is not found. Exiting script" -ForegroundColor Red
		EXIT
	}
}

#Function to delete old files based on age
Function Invoke-Housekeeping {
    [CmdletBinding()] 
    param ( 
        [Parameter(Mandatory)] 
        [string]$folderPath,
    
		[Parameter(Mandatory)]
		[int]$daysToKeep
    )
    
    $datetoDelete = (Get-Date).AddDays(-$daysToKeep)
    $filesToDelete = Get-ChildItem $FolderPath | Where-Object { $_.LastWriteTime -lt $datetoDelete }

    if (($filesToDelete.Count) -gt 0) {	
		foreach ($file in $filesToDelete) {
            Remove-Item -Path ($file.FullName) -Force -ErrorAction SilentlyContinue
		}
	}	
}
#----------------------------------------------------------------------------------------------------
#kill transcript if still running--------------------------------------------------------------------
try{
    stop-transcript|out-null
  }
  catch [System.InvalidOperationException]{}
#----------------------------------------------------------------------------------------------------
#start transcribing----------------------------------------------------------------------------------
if ($enableDebug -eq $true) {Start-Transcript -Path $logFile}
#----------------------------------------------------------------------------------------------------
#BEGIN------------------------------------------------------------------------------------------
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Begin" -ForegroundColor Green
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Connecting to MsOnline Service" -ForegroundColor Green

#Connect to O365 Shell
#Note: This uses an encrypted credential (XML). To store the credential:
#1. Login to the Server/Computer using the account that will be used to run the script/task
#2. Run this "Get-Credential | Export-CliXml msOnlineStoredCredential.xml"
#3. Make sure that msOnlineStoredCredential.xml is in the same folder as the script.
$onLineCredential = Import-Clixml "$($script_root)\msOnlineStoredCredential.xml"
New-MsOLSession $onLineCredential

#Start Export Process---------------------------------------------------------------------------
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Getting SKUs" -ForegroundColor Yellow
$tenantSku = Get-MsolAccountSku
$tenantName = (Get-MsolCompanyInformation).DisplayName
$skus = $skus | Where-Object {$_.Monitor -eq 1}

foreach ($sku in $skus) {
	
	$x = $tenantSku | Where-Object {$_.SkuPartNumber -eq "$($sku.SkuPartNumber)"}
	if ($x.count -gt 0)	{
		$temp = "" | Select-Object Sku,License,Total,Consumed,Available,WarningThreshold,Status
		$temp.License = $sku.SkuFriendlyName
		$temp.Sku = $Sku.SkuPartNumber
		$temp.Total = $x.ActiveUnits
		$temp.Consumed = $x.ConsumedUnits
		$temp.Available = $x.ActiveUnits - $x.ConsumedUnits
		if ($temp.Available -lt 1) {$temp.Available = 0}
		$temp.WarningThreshold = $sku.Threshold
		if ($temp.Available -gt $sku.Threshold) {$temp.Status = "Good"}
		if ($temp.Available -le $sku.Threshold) {$temp.Status = "Warning"}
		if ($sku.Threshold -eq 0) {$temp.Status = "Not Set"}
		$temp | Export-Csv $reportFile -NoTypeInformation -Append
	}
}

$items = Import-Csv $reportFile | Sort-Object Status -Descending
$items | Format-Table -AutoSize
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": CSV Saved to $reportFile" -ForegroundColor Yellow

#Invoke Housekeeping----------------------------------------------------------------------------
if ($enableHousekeeping -eq $true){
	Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Deleting reports older than $($daysToKeep) days" -ForegroundColor Yellow
	Invoke-Housekeeping -folderPath $reportPath -daysToKeep $daysToKeep
}
#-----------------------------------------------------------------------------------------------
#$timeTaken = New-TimeSpan -Start $Today -End (Get-Date)

#Send email if option is enabled ---------------------------------------------------------------
if ($SendReport -eq $true){	

$xSubject="[$($tenantName)] $($mailSubject): " + ('{0:dd-MMM-yyyy hh:mm:ss tt}' -f $Today)
$htmlBody=@'
<!DOCTYPE html>
<html>
<head>
<style>
table {
  font-family: "Century Gothic", sans-serif;
  border-collapse: collapse;
  width: 100%;
}
td, th {
  border: 1px solid #dddddd;
  text-align: left;
  padding: 8px;
}
</style>
</head>
<body>
<table>
'@
$htmlBody+="<tr><th>License ID</th><th>License Name</th><th>Total</th><th>Consumed</th><th>Available</th><th>Threshold</th><th>Status</th></tr>"
foreach ($item in $items){
	$htmlBody+="<tr><td>$($item.Sku)</td><td>$($item.License)</td><td>$($item.Total)</td><td>$($item.Consumed)</td><td>$($item.Available)</td><td>$($item.WarningThreshold)</td>"
	if ($item.Status -eq 'Warning') {
		$htmlBody+="<td><b><font color=""red"">$($item.Status)</font></b></td></tr>"
	}
	elseif ($item.Status -eq 'Good') {
		$htmlBody+="<td><b><font color=""Green"">$($item.Status)</font></b></td></tr>"
	}
	else {
		$htmlBody+="<td><b><font color=""Blue"">Not Set</font></b></td></tr>"
	}	
}

$htmlBody+="</table>"
$htmlBody+="<p style=""font-family:Century Gothic;""><a href=""https://github.com/junecastillote/Get-O365SkuReport"" target=""_blank"">Get-O365SkuReport v$($scriptVersion)</a></p>"
$htmlBody+="</body></html>"

$htmlBody | Out-File $htmlFile

Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": HTML Saved to $htmlFile" -ForegroundColor Yellow
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Sending report to" ($toAddress -join ";") -ForegroundColor Yellow
Send-MailMessage -from $fromAddress -to $toAddress.Split(",") -subject $xSubject -body $htmlBody -dno onSuccess, onFailure -smtpServer $SMTPServer -Port $smtpPort -Credential $onLineCredential -UseSsl -BodyAsHtml
}
#-----------------------------------------------------------------------------------------------
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": End" -ForegroundColor Green
#-----------------------------------------------------------------------------------------------
#kill transcript if still running
try{
    stop-transcript|out-null
  }
  catch [System.InvalidOperationException]{}