#抓取要監控的電腦清單
$computers= get-Content ./Computer.txt
#已使用超過%
$WARING='25'

#寄送通知的人
$TO="bchang@cathaylife.com.tw"

foreach ($computer in $computers)
    {
        $Disks = Get-WmiObject -computer $computer -Class Win32_LogicalDisk -Filter "DriveType = 3"  
  
            foreach ($Disk in $Disks)  
            {  
                $ID = "磁碟機：{0}" -f $Disk.DeviceID  
                $Label = "磁碟機名稱：{0}" -f $Disk.VolumeName  
                $Size = "磁碟機大小：{0:0.0} GB" -f ($Disk.Size / 1GB)  
                $FreeSpace = "剩餘的空間：{0:0.0} GB" -f ($Disk.FreeSpace / 1GB)  
                $Used = ([int64]$Disk.size - [int64]$Disk.FreeSpace)  
                $SpaceUsed = "已用的空間：{0:0.0} GB" -f ($Used / 1GB)  
                $Percent = ($Used * 100.0)/$Disk.Size
                $Percent = [int64]$Percent
                $Percents = "已用的比例：{0:N0}" -f $Percent 
    
                #"---------------------"  
                #"$ID"
                #"$Label"  
                #"$Size"  
                #"$FreeSpace"  
                #"$SpaceUsed"  
                #"$Percent %"
    

                if( $Percent -gt 25 )
                    {
                        $subject = "$computer 主機 $ID 已使用 $Percent % ，請處理"
                        $body = 
" $ID
 $Label
 $Size
 $Size
 $FreeSpace
 $SpaceUsed
 $Percents %"
                                  
                        Send-MailMessage -smtpserver sendmail.cathaylife.com.tw -to $TO -from 9100200@cathlife.com.tw -Subject $subject -body $body 
                        #write-host $body
                    }  
        } 
    } 
    
  