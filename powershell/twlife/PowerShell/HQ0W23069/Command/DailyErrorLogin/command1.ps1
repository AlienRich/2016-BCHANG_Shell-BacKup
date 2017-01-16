# 匯入Server 清單
$computers = get-Content D:\Command\computer-list1.txt

# Lets create our variables
$HostInfo = @()

#取得時間
$Date = Get-date

# Fon 1
# 將Server清單匯入
  foreach ($computer in $computers) {
    $z = "<style>"
    $z = $z + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
    $z = $z + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
    $z = $z + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod}"
    $z = $z + "</style>"
     
     # 抓取Server 登入失敗情況
     $user = $computer + "\Administrator"
     $Info = Get-EventLog Security -computer $computer -EntryType FailureAudit -After ($Date).adddays(-1) | Where {$_.message –match "Administrator"} 
     IF ( $Info -eq $null)
     { $Info = "$computer"+"<br>無資料<br>" }
     ELSE{
     
     $Info = $Info |Select EntryType,TimeWritten,Message| ConvertTo-HTML -head $z
     $Info = "$computer" + $Info
     }
     $HostInfo += $Info
    
}
     # Lets output to the console
     $HostInfo | out-file d:\Command\D1.html

#Fon 2
foreach ($computer in $computers) {
    $z = "<style>"
    $z = $z + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
    $z = $z + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
    $z = $z + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod}"
    $z = $z + "</style>"
     
    #抓取Server 登入失敗情況
    $Info2 = Get-EventLog System -computer $computer -After ($Date).adddays(-1) |Where {$_.message –match "重新啟動" -or $_.message -match "關機"}
    IF ( $Info2 -eq $null)
    { $Info2 = "$computer"+"<br>無資料<br>" }
    ELSE{
     
    $Info2 = $Info2 |Select EntryType,TimeWritten,Message| ConvertTo-HTML -head $z
    $Info2 = "$computer" + $Info2
    }
    $HostInfo2 += $Info2
    
}
    #Lets output to the console
    $HostInfo2 | out-file d:\command\D3.html
     
#Function Send-Mail
#Send-MailMessage -smtpserver hq0w23074 -to bchang@twlife.com.tw,jerrywu@twlife.com.tw -from SIM@twlife.com.tw -Subject "CheckList" -Encoding ([System.Text.Encoding]::UTF8) -Attachment "Check.txt"
