#Import-Module sqlps

$serverName = "cxlsvr55"
$DB = "svrmanage"
$uid ="isvrmanage"
$pwd = "5678" 
#$credential = Get-Credential 
#$userName = $credential.UserName.Replace("\","")  
#$pass = $credential.GetNetworkCredential().password  
$insert_stmt = "INSERT INTO CurrDisk (ComputerName, DiskName ,FreeSpace ,TotalSpace ,FreePercent  ) 
                        VALUES ('CXLSVR136', 'C:' ,'1' ,'1' ,'4' )"

invoke-sqlcmd -query $insert_stmt -database $DB -serverinstance $servername -username $uid -password $pwd 
