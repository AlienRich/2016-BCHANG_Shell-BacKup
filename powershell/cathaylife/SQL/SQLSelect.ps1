$SQLServer = "Cxlsvr55"
$SQLDBName = "svrmanage"
$uid ="isvrmanage"
$pwd = "5678"
$SqlQuery = "Select * from CurrDisk "
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $uid; Password = $pwd;"
#$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True; User ID = $uid; Password = $pwd;"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet

#$insert_stmt = "INSERT INTO CurrDisk (ComputerName, DiskName ,FreeSpace ,TotalSpace ,FreePercent ,DateTime ) VALUES ('test', 'C' ,'1.543' ,'19.533' ,'7.9' ,'C2015/8/7 上午 07:00:01' )"
 
## Create your command
 #$SqlCmd.CreateCommand()
# $cmd=$SqlCmd.CommandText = $insert_stmt
 
## Invoke the Insert statement
# $cmd.ExecuteNonQuery()
 
## Close DB Connection
 #$cmd.Close()

$SqlAdapter.Fill($DataSet)

#$DataSet.Tables[0]

