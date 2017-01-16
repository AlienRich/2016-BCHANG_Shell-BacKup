# 匯入Server 清單
$computers = get-Content D:\Command\DailyErrorLogin\computer-list.txt

# Lets create our variables
$HostInfo = @()

# Fon 1
# 將Server清單匯入
  foreach ($computer in $computers) {
     $z = "<style>"
    #$z = $z + "BODY{background-color:peachpuff;}"
    $z = $z + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
    $z = $z + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
    $z = $z + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod}"
    $z = $z + "</style>"
     
     # 抓取Server 登入失敗情況
     $Info = Get-EventLog Security -computer $computer -EntryType FailureAudit -After (get-date).adddays(-1)
     IF ( $Info -eq $null)
     { $Info = "$computer"+"<br>無資料<br>" }
     ELSE{
     
     $Info = $Info |Select EntryType,TimeWritten,Message| ConvertTo-HTML -head $z
     $Info = "$computer" + $Info
     }
     $HostInfo += $Info
    
}
     # Lets output to the console
     $HostInfo | out-file d:\DailyErrorLogin.html

#Fon 2
foreach ($computer in $computers) {
     $z = "<style>"
    #$z = $z + "BODY{background-color:peachpuff;}"
    $z = $z + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
    $z = $z + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
    $z = $z + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod}"
    $z = $z + "</style>"
     
     # 抓取Server 登入失敗情況
     $Info2 = Get-EventLog System -computer $computer -Source USER32 -After (get-date).adddays(-1)
     IF ( $Info2 -eq $null)
     { $Info2 = "$computer"+"<br>無資料<br>" }
     ELSE{
     
     $Info2 = $Info2 |Select EntryType,TimeWritten,Message| ConvertTo-HTML -head $z
     $Info2 = "$computer" + $Info2
     }
     $HostInfo2 += $Info2
    
}
     # Lets output to the console
     $HostInfo2 | out-file d:\DailyHighCommand.html
     
#Function Send-Mail
#Send-MailMessage -smtpserver hq0w23074 -to bchang@twlife.com.tw,jerrywu@twlife.com.tw -from SIM@twlife.com.tw -Subject "CheckList" -Encoding ([System.Text.Encoding]::UTF8) -Attachment "Check.txt"
