#匯入AD模組
Import-module ActiveDirectory

#密碼產生器
$Chars = [Char[]]"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXY0123456789"
$Chars1 = [Char[]]"abcdefghijklmnopqrstuvwxyz"
$Chars2 = [Char[]]"ABCDEFGHIJKLMNOPQRSTUVWXY"
$Chars3 = [Char[]]"0123456789"
$Password1 = ($Chars | Get-Random -count 5 ) -join ""
$Password2 = ($Chars1 | Get-Random -count 1 ) -join ""
$Password3 = ($Chars2 | Get-Random -count 1 ) -join ""
$Password4 = ($Chars3 | Get-Random -count 1 ) -join ""
$Password = $Password1+$Password2+$Password3+$Password4

#修改的帳號
Set-ADAccountPassword -Identity Vision_Tiptop -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force)

#寄送密碼
Send-MailMessage -smtpserver 10.194.99.207 -to maxchung@pegavision.com,alexwu@pegavision.com,TomKao@pegavision.com -from TipTopPassWord@pegavision.com -Subject "TipTopPassWord" -Body $Password