#匯入SQL模組
Import-Module sqlps -DisableNameChecking

#SQLServer相關參數　
$SQLServer = "Cxlsvr55"
$SQLDBName = "svrmanage"
$uid ="isvrmanage"
$pwd = "5678"

#掃描匯入電腦
$Computers = get-Content D:\PROGRAM\PowerShell\03.Query_Disk_To_Sql\DelList.txt

    foreach ($computer in $Computers)
        {
        #刪除作業
        $DelTable = "DELETE FROM CurrDisk WHERE Computername like '$Computer'"
        invoke-sqlcmd -query $Deltable -database $SQLDBName -serverinstance $SQLServer -username $uid -password $pwd
        }