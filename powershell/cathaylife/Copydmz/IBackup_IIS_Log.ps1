# �פJ�D���M��
$ServerList = Import-csv D:\PROGRAM\9002200\copydmzlog\ServerList.csv

Foreach($Server in $ServerList)
{
    $Source = $Server.Source
    $destination = $Server.destination
    #�s�u�����Ϻо�
    net use i: "$Source" dmz9002200 /user:cxlsna\i100000
    #LOG��m
    $Log = "D:\PROGRAM\9002200\copydmzlog\copydmzlog.txt"

    #�}�l��
    get-childitem $source -recurse | Where {$_.LastWriteTime.Date -eq ((get-date).adddays(-1).date)} |foreach {
    copy $_.fullname $destination -Force -Recurse -errorAction silentlyContinue
    if($? -eq $false){echo "$($_.fullname) �ƻs�ɮץ��� to $destination" | out-file -append $log} 
    else 
    {write-output "$($_.fullname) �ƻs�ɮק��� to $destination"  | out-file -append $log}
    }

    #���������Ϻо�
    net use i: /DELETE /y
    #Write-Host $Source
    #Write-Host $destination
    
}

