$StartDate = Get-Date
$File = "C:\ERP.txt"

net use x: \\10.194.112.100\erp_backup$\erp /user:pegavision\sp02 Bchang00
cd x:\
ftp -s:tiptop_backup.txt >> $file
$StopDate = Get-Date

Send-MailMessage -smtpserver 10.194.99.208 -to billchang@pegavision.com,alexwu@pegavision.com,maxchung@pegavision.com -from ServerBackup@pegavision.com -Subject "ERP Backup LOG" -Body "備份起始時間 $StartDate `n備份結束時間 $StopDate" -Encoding ([System.Text.Encoding]::UTF8) -Attachment $file
Remove-Item $File

