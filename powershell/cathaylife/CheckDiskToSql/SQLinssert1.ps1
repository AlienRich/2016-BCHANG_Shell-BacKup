#Import-Module sqlps

$serverName = "cxlsvr55"
$DB = "svrmanage"
$uid ="isvrmanage"
$pwd = "5678" 
#$credential = Get-Credential 
#$userName = $credential.UserName.Replace("\","")  
#$pass = $credential.GetNetworkCredential().password  
$insert_stmt = "DELETE FROM CurrDisk WHERE Computername like 'CXLSVR136'"

invoke-sqlcmd -query $insert_stmt -database $DB -serverinstance $servername -username $uid -password $pwd 
