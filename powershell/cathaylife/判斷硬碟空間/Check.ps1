#抓取硬碟D槽容量
$Disk=Get-WmiObject win32_logicaldisk -Filter "DeviceID = 'D:' " | select-object DeviceID,@{Name="UseSpace";Expression={100-$_.FreeSpace/$_.Size*100}}

#抓取硬碟D槽容量剩餘%
$UseSpace = $Disk | select UseSpace

#去除不要的字
$b=$UseSpace -replace "@{UseSpace=", ""
$b=$b -replace "}", ""
#取整數
$a =[int]$b

#如果大於75%救寄送信件
if( $a -gt 75 ){
Send-MailMessage -smtpserver hq0w23074 -to bchang@cathaylife.com.tw -from CXLSVR239@cathaylife.com.tw -Subject "CXLSVR239硬碟空間超過90%，請處理" 
#write-host 大於75
}


