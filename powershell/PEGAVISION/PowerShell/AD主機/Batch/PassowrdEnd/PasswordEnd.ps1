Import-Module Activedirectory
$alladuser=get-aduser -filter *  | %{$_.Samaccountname}
$userlist = @()
#################################################
#檢測AD密碼過期時間與郵件通知帳戶
##################################################
foreach ($user in $alladuser){
#抓取最後一次修改密碼時間
$pwdlastset=Get-ADUser $user -Properties * | %{$_.passwordlastset}
#抓取顯示名稱
$DisplayName=Get-ADUser $user -Properties * | %{$_.DisplayName}
#抓取Mail
$MailTo=Get-ADUser $user -Properties * | %{$_.mail}
#密碼過期時間
$pwdlastday=($pwdlastset).adddays(90)
#當下時間
$now=get-date
#判斷帳戶是否設定永久有效
$neverexpire=get-aduser $user -Properties * |%{$_.PasswordNeverExpires}
#距離密碼過期時間
$expire_days=($pwdlastday - $now).Days
#判斷過期時間小於15天的且沒有定永久有效的帳戶
if($expire_days -lt 7 -and $neverexpire -like "false" ){
    $chineseusername= Get-ADUser $user  -Properties * | %{$_.Displayname}
    #邮件正文
    $Emailbody=
    "Dear $DisplayName :
    您的郵件帳號密碼即将在 $expire_days 天候過期， $pwdlastday 之後將無法使用，請盡快更改。
    密碼到期未變更，將導致無法使用公司內部網路，請依規定進行密碼變更，謝謝配合
    請勿回復此郵件, 謝謝配合!
    
    重置密碼原則：
    ○密碼長度最少 8 位；
    ○不能使用之前最近使用的 3 個密碼
    ○密碼必須使用複雜需求（大寫、小寫、數字和符號中必須含3種）
"
Send-MailMessage -from it@pegavision.com -to $MailTo -subject "[系統通知]您的Email帳號密碼將於 $expire_days 天後過期,請立即更改密碼" -body $Emailbody -Attachments "C:\Batch\PassowrdEnd\修改MAIL密碼方式.pdf" -smtpserver mail.pegavision.com -Encoding ([System.Text.Encoding]::UTF8)
#############################################
#查找账户的密码过期时间并发送至管理员账户
#############################################
$username=Get-ADUser $user  -Properties *
$userobject=New-object psobject
$userobject | Add-Member -membertype noteproperty -Name 名稱            -value $username.displayname
$userobject | Add-Member -membertype noteproperty -Name 郵件              -Value $username.mail
$userobject | Add-Member -membertype noteproperty -Name 最後一次修改密碼時間  -Value $username.Passwordlastset
$userobject | Add-Member -membertype noteproperty -Name 密碼過期時間      -Value $pwdlastday
$userobject | Add-Member -membertype noteproperty -Name 距離到期天數  -Value $expire_days
$userlist+=$userobject
}
}
$EmailbodyHTML=$userlist|
sort-object 距離到期天數 |
ConvertTo-Html |
Out-String
Send-Mailmessage -from  it@Pegavision.com –to maxchung@pegavision.com,billchang@pegavision.com -Bodyashtml $EmailbodyHTML -Subject 郵件管理員通知 -smtpserver mail.pegavision.com -Encoding ([System.Text.Encoding]::UTF8)