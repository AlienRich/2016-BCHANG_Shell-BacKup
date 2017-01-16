$StartDate = Get-Date
$Computer = hostname
$DelFile = "C:\Delfile.txt"
$BakFile = "C:\Backup.txt"
$MailTo = "billchang@pegavision.com"

net use x: \\10.194.112.100\serverbackup$\$Computer /user:pegavision\sp02 Bchang00
#刪除資料
$file1 = Get-ChildItem -Path x:\ARES -Force | Remove-Item -force -recurse
$file1 = Get-ChildItem -Path x:\DB_Dump -Force | Remove-Item -force -recurse
$file1 = Get-ChildItem -Path x:\HCP_SRC -Force | Remove-Item -force -recurse

#備份今天產生的資料
robocopy D:\ARES X:\ARES /mir >>$BakFile
robocopy D:\DB_Dump X:\DB_Dump /mir >>$BakFile
robocopy D:\HCP_SRC X:\HCP_SRC /mir >>$BakFile
#$file2 = get-item D:\ARES
#$file2 >> $BakFile
#$file2 | Copy-Item -destination x:\

$StopDate = Get-Date

Send-MailMessage -smtpserver 10.194.99.208 -to billchang@pegavision.com,maxchung@pegavision.com -from ServerBackup@pegavision.com -Subject "$Computer Backup LOG" -Body "備份起始時間 $StartDate `n備份結束時間 $StopDate" -Encoding ([System.Text.Encoding]::UTF8) -Attachment $DelFile,$BakFile
Remove-Item $DelFile,$BakFile
net use x: /delete
