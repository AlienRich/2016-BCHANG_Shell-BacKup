$filename = ¡§C:\CCRCheck.html¡¨
$smtpServer = ¡§hq0w23074.twlife.com.tw¡¨

$msg = new-object Net.Mail.MailMessage
$att = new-object Net.Mail.Attachment($filename)
$smtp = new-object Net.Mail.SmtpClient($smtpServer)

$msg.From = ¡§hptest@twlife.com.tw¡¨
$msg.To.Add(¡§bchang@twlife.com.tw¡¨)
$msg.Subject = ¡§CCR Status¡¨
$msg.Body = ¡§C:\CCRCheck.log¡¨
$msg.Attachments.Add($att)

$smtp.Send($msg)
