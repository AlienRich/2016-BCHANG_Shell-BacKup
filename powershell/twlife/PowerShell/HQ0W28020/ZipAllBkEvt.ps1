#Function �ܼ�
$FormDate=get-date -Format yyyyMMdd
$SelectFile=Get-ChildItem -Path D:\EventLog\Event\


#Function �|�XLOG�����Y
get-date -Format hhmm | out-file D:\Batch\Event_$FormDate.txt -encoding Unicode
echo "�ƥ��}�l" | out-file  D:\Batch\Event_$FormDate.txt -append 
$SelectFile | Select-Object FullName,Length | out-file D:\Batch\Event_$FormDate.txt -append
cd "c:\Program Files (x86)\7-Zip\"
.\7z.exe a -tzip D:\EventLog\$FormDate.zip D:\EventLog\Event
get-date -Format hhmm |out-file D:\Batch\Event_$FormDate.txt -append 
echo "�ƥ�����`r`n" | out-file D:\Batch\Event_$FormDate.txt -append 

#Function �ܼ�
$log=Get-Content D:\Batch\Event_$FormDate.txt

#Function Send-Mail
Send-MailMessage -smtpserver hq0w23074 -to bchang@twlife.com.tw -from FileServer@twlife.com.tw -Subject "FileServer-HQ-Event" -Encoding ([System.Text.Encoding]::UTF8) -body "$log"

#Function �R��
Remove-Item D:\Batch\Event_$FormDate.txt
Remove-Item D:\EventLog\Event\*.*