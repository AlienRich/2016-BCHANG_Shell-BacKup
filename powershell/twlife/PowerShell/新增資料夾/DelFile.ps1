#Function 變數
$FormDate=get-date -Format yyyyMMdd
$SelectFile=Get-ChildItem -Path "E:\temp" -Include *.* -Recurse | where {$_.lastwritetime -lt (get-date).adddays(-3)}

#Function Del old data
Remove-Item e:\test\*.*

#Function 產生LOG
$SelectFile | Select-Object FullName,CreationTime | Export-Csv e:\$FormDate.csv

#Function Move檔案
Move-Item $SelectFile e:\test

#Function Send-Mail&Delect Csv
Send-MailMessage -smtpserver hq0w23074 -to bchang@twlife.com.tw -from test@twlife.com.tw -Subject "mail test" -Attachments e:\$FormDate.csv 

Remove-Item e:\$FormDate.csv  
