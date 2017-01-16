$StartDate = Get-Date
$Computer = hostname
$DelFile = "c:\Delfile.txt"
$BakFile = "c:\Backup.txt"
#$MailTo = "billchang@pegavision.com"

net use x: \\10.194.112.100\serverbackup$\$Computer\Daily /user:pegavision\sp02 Bchang00
#刪除資料夾超過30天
$file1 = Get-ChildItem -Path x:\ -Force | Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-365)}
$file1 >> $DelFile
$file1 | Remove-Item

#備份今天產生的資料
$file2 = Get-ChildItem -Path "D:\DB_Dump" -Force | Where-Object {$_.LastWriteTime -gt (Get-Date).AddDays(-2)}
$file2 >> $BakFile
$file2 | Copy-Item -destination x:\

$StopDate = Get-Date

Send-MailMessage -smtpserver 10.194.99.208 -to billchang@pegavision.com,maxchung@pegavision.com -from ServerBackup@pegavision.com -Subject "$Computer Backup LOG" -Body "備份起始時間 $StartDate `n備份結束時間 $StopDate" -Encoding ([System.Text.Encoding]::UTF8) -Attachment $DelFile,$BakFile
Remove-Item $DelFile,$BakFile
net use x: /delete
