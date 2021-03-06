#匯入SQL模組
Import-Module sqlps -DisableNameChecking

#SQLServer相關參數　
$SQLServer = "Cxlsvr55"
$SQLDBName = "svrmanage"
$uid ="isvrmanage"
$pwd = "5678"

#掃描匯入電腦
$Computers = get-Content D:\PROGRAM\PowerShell\03.Query_Disk_To_Sql\QueryList.txt

    foreach ($computer in $Computers)
        {
        $Disks = Get-WmiObject -computer $computer -Class Win32_LogicalDisk -Filter "DriveType = 3"  
        
        #查詢作業
        $Query = invoke-sqlcmd -query "select computername from CurrDisk where computername like '$Computer'" -database $SQLDBName -serverinstance $SQLServer -username $uid -password $pwd

        If ($Query)
              { foreach ($Disk in $Disks)  
                 {  
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
                
                 #----------資料庫作業-----------------
                 #Update 作業
                 $UpdateTable= "Update CurrDisk
                 set DateTime ='$Date',FreePercent ='$PercentFree',FreeSpace='$FreeSpace'
                 where ComputerName = '$Computer'
                 and DiskName = '$ID'"
                                                              
                 invoke-sqlcmd -query $UpdateTable -database $SQLDBName -serverinstance $SQLServer -username $uid -password $pwd
                 }
               }
                    else
                        {  foreach ($Disk in $Disks)  
                            {  
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
                
                            #----------資料庫作業-----------------
                            #新增作業
                            $InsertTable = "INSERT INTO CurrDisk (ComputerName, DiskName ,FreeSpace ,TotalSpace ,FreePercent,DateTime  ) 
                                VALUES ('$Computer', '$ID' ,'$FreeSpace' ,'$Size' ,'$PercentFree','$Date'  )"
                            invoke-sqlcmd -query $InsertTable -database $SQLDBName -serverinstance $SQLServer -username $uid -password $pwd
                            }
                        }
      }


      #set FreePercent =' $PercentFree'
      #set FreeSpace='$FreeSpace'