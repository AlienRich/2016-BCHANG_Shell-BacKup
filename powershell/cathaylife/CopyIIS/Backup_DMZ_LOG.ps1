# 匯入主機清單
$ServerList = Import-csv D:\PROGRAM\9002200\copy_txtlog\ServerList.csv

Foreach($Server in $ServerList)
{
    $IP = $Server.IP
    $ServerName = $Server.Servername
    #連線網路磁碟機
    net use i: \\$IP\d$ dmz9002200 /user:cxlsna\i100000
    #抓取要搬移的檔案資訊
    $file = Get-ChildItem i:\iislog\eventlog\Security_TXT\SVRQRY*.txt | ForEach-Object{$_.Name}
    #搬移的來源端
    $Source = "i:\iislog\eventlog\Security_TXT\SVRQRY*.txt"
    #搬移的目的端
    $destination = "\\cxlsvr184\e$\EventLog\Security_TXT\$ServerName\" 
    #LOG位置
    $Log = "D:\PROGRAM\9002200\copy_txtlog\$ServerName.txt"

    #開始備份
    "===============================" >>$Log
    get-date -Format "yyyy-MM-dd hh:mm 開始備份" >>$Log
    $file >>$Log
    move $Source $destination 
    get-date -Format "yyyy-MM-dd hh:mm 結束備份" >>$Log
    "===============================" >>$Log

    #移除網路磁碟機
    net use i: /DELETE /y
    
}

