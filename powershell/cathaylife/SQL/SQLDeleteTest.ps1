
    $SQLServer = "Cxlsvr55"
                $SQLDBName = "svrmanage"
                $uid ="isvrmanage"
                $pwd = "5678"
                $conn = New-Object System.Data.SqlClient.SqlConnection
                $conn.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $uid; Password = $pwd;"
                Try
                {
                    ## Open DB Connection
                    $conn.Open()
                    ## Create your command
                    $cmd = $conn.CreateCommand()
                    $insert_stmt1 = "DELETE FROM CurrDisk WHERE Computer (cxlsvr136)"
                    
                    #Create SQL Insert Statement with your values
                    #$insert_stmt = "INSERT INTO CurrDisk (ComputerName, DiskName ,FreeSpace ,TotalSpace ,FreePercent,DateTime  ) 
                    #    VALUES ('CXLSVR136', '$ID' ,'$FreeSpace' ,'$Size' ,'$PercentFree','$Date'  )"
                    $cmd.CommandText = $insert_stmt1
                    ## Invoke the Insert statement
                    $cmd.ExecuteNonQuery()
                }
            finally
                {
                    if ($conn -and ($conn.state -eq 'Open'))
                {
                    $conn.Close()
                }
                }