$StartDate = Get-Date
$Computer = hostname
$File = "C:\$computer.txt"
$MailTo = "billchang@pegavision.com"

net use x: \\10.194.112.100\serverbackup$\$Computer /user:pegavision\sp02 Bchang00
wbadmin start backup -backupTarget:\\10.194.112.100\serverbackup$\$Computer -include:D: -allcritical -quiet >> $File

$StopDate = Get-Date

Send-MailMessage -smtpserver 10.194.99.208 -to billchang@pegavision.com,maxchung@pegavision.com -from ServerBackup@pegavision.com -Subject "$Computer Backup LOG" -Body "備份起始時間 $StartDate `n備份結束時間 $StopDate" -Encoding ([System.Text.Encoding]::UTF8) -Attachment $file
Remove-Item $File
net use x: /delete
