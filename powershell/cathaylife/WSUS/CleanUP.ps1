#匯入清單
$CleanUPLists = Get-Content D:\Batch\CleanUPList.txt

# Variablen
$DateFormat = Get-Date -format yyyyMMdd-HH-mm
$Logfile = "D:\CleanLog\Wsus-Cleanup-$DateFormat.log"

    foreach ($CleanUPList in $CleanUPLists)
    {
    
    "$CleanUPList 清理作業" >> $Logfile
    # WSUS Bereinigung durchfuhren
    Get-WsusServer $CleanUPList -PortNumber 80 | Invoke-WsusServerCleanup -CleanupObsoleteUpdates -CleanupUnneededContentFiles -CompressUpdates -DeclineExpiredUpdates -DeclineSupersededUpdates >> $Logfile
    " " >> $Logfile   
    }

# Mail Variablen
$MailSMTPServer = "cathaymail.linyuan.com.tw"
$MailFrom = "WSUSAdmin@Cathaylife.com.tw"
$MailTo = "Bchang@cathaylife.com.tw,shaoping@cathaylife.com.tw"
$MailSubject = "WSUS清理作業紀錄"
$MailBody = Get-Content $Logfile | Out-String

# Mail versenden
Send-MailMessage -SmtpServer $MailSMTPServer -From $MailFrom -To $MailTo -subject $MailSubject -body $MailBody -Encoding Unicode