#抓取要監控的電腦清單
$Lists= Import-CSV D:\PROGRAM\PowerShell\01.Check_Disk\List.csv
#已使用超過%
$WARING='90'

#寄送通知的人
#$TO="bchang@cathaylife.com.tw"


foreach ($List in $Lists)
    {
        $computer = $List.Server
        $Mail = $List.to
        [string[]]$To = $Mail.Split(',')

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
    

                if( $Percent -gt $WARING )
                    {
                        $subject = "$computer 主機 $ID 已使用 $Percent % ，請處理"
                        $body = 
" $ID
 $Size
 $FreeSpace
 $SpaceUsed
 $Percents %"
                                  
                        Send-MailMessage -smtpserver sendmail.cathaylife.com.tw -to $To -from "系統資訊一科-系統資訊一科 <9100200@cathlife.com.tw>" -Subject $subject -body $body -Encoding ([System.Text.Encoding]::UTF8) 
                        #write-host $body
                    }  
        } 
    } 
    
  