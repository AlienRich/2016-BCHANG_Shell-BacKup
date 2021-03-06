Import-Module Activedirectory
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010

$Results = ForEach ($User in (Get-ADUser -SearchBase "OU=NTCC,DC=taiwancement,DC=com" -Filter * -Properties Company))
{   $Mail = Get-Mailbox $User.Name -ErrorAction SilentlyContinue | Get-MailboxStatistics | where {$_.lastlogontime -lt (get-date).adddays(-7)}
    If ($Mail)
    {   New-Object PSObject -Property @{
            Name = $User.Name
            Company = $User.Company
            lastlogontime = $Mail.lastlogontime 
        }
    }
}
$Results | Select Company,Name,lastlogontime | export-csv -delimiter "`t" -path C:\Batch\OwaLogin\OwaLogin.txt -Encoding UTF8