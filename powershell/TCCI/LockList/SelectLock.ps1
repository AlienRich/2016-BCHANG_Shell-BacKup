$logName = "security"
$pcName = "TCC-HQAD04", "TCCGROUP-HQAD", "TCCGROUP-HQAD03", "TCCGROUP-HQAD05", "TCCGROUP-HQAD06"
$eventID = "4740"
Get-EventLog -LogName $logName -ComputerName $pcName | where {$_.eventID -eq $eventID} `
| fl -Property timegenerated, replacementstrings, message >> C:\Batch\LockList\Log.txt 

Send-MailMessage -SmtpServer TCCGROUP-HT01 -To bill.chang@tcci.com.tw -From LockList@tcci.com.tw -Subject "LockList" -Attachment C:\Batch\LockList\Log.txt

Del C:\Batch\LockList\Log.txt