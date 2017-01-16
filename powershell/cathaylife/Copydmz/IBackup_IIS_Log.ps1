# 匯入主機清單
$ServerList = Import-csv D:\PROGRAM\9002200\copydmzlog\ServerList.csv

Foreach($Server in $ServerList)
{
    $Source = $Server.Source
    $destination = $Server.destination
    #連線網路磁碟機
    net use i: "$Source" dmz9002200 /user:cxlsna\i100000
    #LOG位置
    $Log = "D:\PROGRAM\9002200\copydmzlog\copydmzlog.txt"

    #開始備
    get-childitem $source -recurse | Where {$_.LastWriteTime.Date -eq ((get-date).adddays(-1).date)} |foreach {
    copy $_.fullname $destination -Force -Recurse -errorAction silentlyContinue
    if($? -eq $false){echo "$($_.fullname) 複製檔案失敗 to $destination" | out-file -append $log} 
    else 
    {write-output "$($_.fullname) 複製檔案完成 to $destination"  | out-file -append $log}
    }

    #移除網路磁碟機
    net use i: /DELETE /y
    #Write-Host $Source
    #Write-Host $destination
    
}

