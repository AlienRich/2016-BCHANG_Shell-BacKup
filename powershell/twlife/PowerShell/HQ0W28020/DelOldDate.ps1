#Function Del old data
Remove-Item -Path D:\�ӤH��Ƨ�\OldData\*

#Function �ܼ�
$FormDate=get-date -Format yyyyMMdd
$SelectFile=Get-ChildItem -Path "D:\�ӤH��Ƨ�" -Recurse | where {$_.LastWriteTime -lt (get-date).adddays(-7)}  | Where {!$_.PSIsContainer}


#Function ����LOG
$SelectFile | Select-Object FullName,CreationTime | Export-Csv D:\Batch\DelOldDate_$FormDate.csv -NoTypeInformation -encoding "unicode"

#Function Move�ɮ�
$SelectFile | Move-Item -destination D:\�ӤH��Ƨ�\OldData\ 

#Function Send-Mail&Delect Csv
Send-MailMessage -smtpserver hq0w23074 -to bchang@twlife.com.tw -from FileServer@twlife.com.tw -Subject "FileServer-HQ-DelOldDate" -Attachments D:\Batch\DelOldDate_$FormDate.csv 

Remove-Item -Path D:\Batch\DelOldDate_$FormDate.csv  
