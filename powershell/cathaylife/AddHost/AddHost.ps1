
$Computers = get-Content D:\個人資料\09.工具區\powershell\cathaylife\AddHost\Computer.txt


foreach ($computer in $Computers)
    {
    
    $hostfile = "\\$computer\c$\Windows\System32\drivers\etc\hosts"
    
    Copy-Item \\$computer\c$\Windows\System32\drivers\etc\hosts  \\$computer\c$\Windows\System32\drivers\etc\hosts_Bak
    "`r" | Out-File $hostfile -encoding ASCII -append
    "192.168.9.9`tcathaymail.linyuan.com.tw" | Out-File $hostfile -encoding ASCII -append
    
    Copy-Item \\$computer\c$\Windows\System32\drivers\etc\hosts_bak D:\個人資料\09.工具區\powershell\cathaylife\AddHost\hosts_bak_$computer

 }