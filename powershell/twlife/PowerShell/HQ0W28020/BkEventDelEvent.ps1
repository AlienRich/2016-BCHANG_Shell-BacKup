#Function 變數
$FormDate=get-date -Format yyyyMMddHH

#Function 將Event匯出至EXECL
get-eventlog security | export-csv -path d:\EventLog\Event\$FormDate.csv -NoTypeInformation -encoding "unicode"

#Function 將Event刪除
clear-eventlog security