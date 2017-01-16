# �פJServer �M��
$computers = get-Content D:\Command\computer-list.txt

# Lets create our variables
$HostInfo = @()

#���o�ɶ�
$Date = Get-date
$Date1 = Get-Date -format yyyyMMdd

# Function 1 ����n�JAdministrator���ѱ��p
# �NServer�M��פJ
  foreach ($computer in $computers) {
    #���ͮج[���C��
    $z = "<style>"
    $z = $z + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
    $z = $z + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
    $z = $z + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod}"
    $z = $z + "</style>"
     
     # ���Server �n�JAdministrator���Ѫ�Log
     $Info = Get-EventLog Security -computer $computer -EntryType FailureAudit -After ($Date).adddays(-1) | Where {$_.message �Vmatch "Administrator"} 
     IF ( $Info -eq $null)
        { $Info = "$computer"+"<br>�L���<br>" }
      ELSE{
            $Info = $Info |Select EntryType,TimeWritten,Message| ConvertTo-HTML -head $z
            $Info = "$computer" + $Info
          }
          $HostInfo += $Info
    
        }
     # ���͸��
     $HostInfo | out-file d:\Command\D2.html

#Function 2 ���Server System Log �� ���s�Ұʩ������� Log
# �NServer�M��פJ
  foreach ($computer in $computers) {
    #���ͮج[���C��
    $z = "<style>"
    $z = $z + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
    $z = $z + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
    $z = $z + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod}"
    $z = $z + "</style>"
     
    #���Server System Log �� ���s�Ұʩ������� Log
    $Info2 = Get-EventLog System -computer $computer -After ($Date).adddays(-1) |Where {$_.message �Vmatch "���s�Ұ�" -or $_.message -match "����"}
    IF ( $Info2 -eq $null)
    { $Info2 = "$computer"+"<br>�L���<br>" }
    ELSE{
     
    $Info2 = $Info2 |Select EntryType,TimeWritten,Message| ConvertTo-HTML -head $z
    $Info2 = "$computer" + $Info2
    }
    $HostInfo2 += $Info2
    
}
    # ���͸��
    $HostInfo2 | out-file d:\Command\D4.html
    
    #�����ƦX��
    $Str1 =  Get-Content 'D:\Command\D1.html'
    $Str2 =  Get-Content 'D:\Command\D2.html'
    $Str3 =  Get-Content 'D:\Command\D3.html'
    $Str4 =  Get-Content 'D:\Command\D4.html'
    
    $strA =  $Str1 + $Str2 
    $strB =  $Str3 + $Str4  
    
    $file1= "d:\Command\ErrorLogin_" + $Date1 + ".html"
    $file2= "d:\Command\HighCommand_" + $Date1 + ".html"
     
    $strA | out-file $file1
    $strB | out-file $file2
    
     
#Function 3 Send-Mail
Send-MailMessage -smtpserver hq0w23074 -to bchang@twlife.com.tw,jerrywu@twlife.com.tw -from SIM@twlife.com.tw -Subject "LOG"  -Attachments $file1,$file2

#Function 4 �ƥ�Log��
copy $file1 "\\hq0w23065\LogBackup\ErrorLog\"
copy $file2 "\\hq0w23065\LogBackup\CommandLog\"

#�M�����
Remove-Item D:\Command\*.html
