#Function 變數
$FormDate=get-date -Format yyyyMMdd
$SelectFile=Get-ChildItem -Path D:\EventLog\Event\


#Function 會出LOG及壓縮
get-date -Format hhmm | out-file D:\Batch\Event_$FormDate.txt -encoding Unicode
echo "備份開始" | out-file  D:\Batch\Event_$FormDate.txt -append 
$SelectFile | Select-Object FullName,Length | out-file D:\Batch\Event_$FormDate.txt -append
cd "c:\Program Files (x86)\7-Zip\"
.\7z.exe a -tzip D:\EventLog\$FormDate.zip D:\EventLog\Event
get-date -Format hhmm |out-file D:\Batch\Event_$FormDate.txt -append 
echo "備份結束`r`n" | out-file D:\Batch\Event_$FormDate.txt -append 

#Function 變數
$log=Get-Content D:\Batch\Event_$FormDate.txt

#Function Send-Mail
Send-MailMessage -smtpserver hq0w23074 -to bchang@twlife.com.tw -from FileServer@twlife.com.tw -Subject "FileServer-HQ-Event" -Encoding ([System.Text.Encoding]::UTF8) -body "$log"

#Function 刪除
Remove-Item D:\Batch\Event_$FormDate.txt
Remove-Item D:\EventLog\Event\*.*