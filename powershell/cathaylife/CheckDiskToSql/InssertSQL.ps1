
$Disks = Get-WmiObject -computer CXLSVR136 -Class Win32_LogicalDisk -Filter "DriveType = 3"  
  
            foreach ($Disk in $Disks)  
            {  
                #$ID = "磁碟機：{0}" -f $Disk.DeviceID  
                #$Label = "磁碟機名稱：{0}" -f $Disk.VolumeName  
                #$Size = "磁碟機大小：{0:0.0} GB" -f ($Disk.Size / 1GB)  
                #$FreeSpace = "剩餘的空間：{0:0.0} GB" -f ($Disk.FreeSpace / 1GB)  
                #$Used = ([int64]$Disk.size - [int64]$Disk.FreeSpace)  
                #$SpaceUsed = "已用的空間：{0:0.0} GB" -f ($Used / 1GB)  
                #$Percent = ($Used * 100.0)/$Disk.Size
                #$Percent = [int64]$Percent
                #$Percents = "已用的比例：{0:N0}" -f $Percent
                $Date = get-date -Format "yyyy/MM/dd HH:mm:ss"
                $ID =  $Disk.DeviceID  
                $Label = $Disk.VolumeName  
                $Size = $Disk.Size / 1GB
                $Size = "{0:N3}" -f $Size
                $FreeSpace = $Disk.FreeSpace / 1GB
                $FreeSpace = "{0:N3}" -f $FreeSpace
                $Used = ([int64]$Disk.size - [int64]$Disk.FreeSpace) 
                $SpaceUsed = $Used / 1GB  
                $Percent = ($Used * 100.0)/$Disk.Size
                $PercentFree = 100 - $Percent
                $PercentFree = "{0:N3}" -f $PercentFree
                #$Percent = [int64]$Percent
                #$Percents = "已用的比例：{0:N0}" -f $Percent 
                
                
                #----------資料庫作業-----------------
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
                    #Create SQL Insert Statement with your values
                    $insert_stmt = "INSERT INTO CurrDisk (ComputerName, DiskName ,FreeSpace ,TotalSpace ,FreePercent,DateTime  ) 
                        VALUES ('CXLSVR136', '$ID' ,'$FreeSpace' ,'$Size' ,'$PercentFree','$Date'  )"
                    $cmd.CommandText = $insert_stmt
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
                }