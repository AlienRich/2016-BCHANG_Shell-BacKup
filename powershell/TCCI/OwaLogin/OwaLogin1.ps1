Import-Module Activedirectory
#Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
Get-ADUser -SearchBase "OU=NTCC,DC=taiwancement,DC=com" -Filter * | foreach-object {
    Get-MailboxStatistics $_.SamAccountName | where {$_.lastlogontime -lt (get-date).adddays(-7)} | Select displayname,lastlogontime
    } | export-csv -delimiter "`t" -path C:\Batch\OwaLogin\OwaLogin.txt -Encoding UTF8
     
Send-MailMessage -SmtpServer TCCGROUP-HT02 -To bill.chang@tcci.com.tw,hank.young@tcci.com.tw -From OwaLogin@tcci.com.tw -Subject "NtccOwaLastLogin" -Attachment C:\Batch\OwaLogin\OwaLogin.txt

#Del C:\Batch\OwaLogin\OwaLogin.txt