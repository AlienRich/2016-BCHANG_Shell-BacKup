#Function Del old data
Remove-Item -Path D:\個人資料夾\OldData\*

#Function 變數
$FormDate=get-date -Format yyyyMMdd
$SelectFile=Get-ChildItem -Path "D:\個人資料夾" -Recurse | where {$_.LastWriteTime -lt (get-date).adddays(-7)}  | Where {!$_.PSIsContainer}


#Function 產生LOG
$SelectFile | Select-Object FullName,CreationTime | Export-Csv D:\Batch\DelOldDate_$FormDate.csv -NoTypeInformation -encoding "unicode"

#Function Move檔案
$SelectFile | Move-Item -destination D:\個人資料夾\OldData\ 

#Function Send-Mail&Delect Csv
Send-MailMessage -smtpserver hq0w23074 -to bchang@twlife.com.tw -from FileServer@twlife.com.tw -Subject "FileServer-HQ-DelOldDate" -Attachments D:\Batch\DelOldDate_$FormDate.csv 

Remove-Item -Path D:\Batch\DelOldDate_$FormDate.csv  
