#Function �ܼ�
$FormDate=get-date -Format yyyyMMddHH

#Function �NEvent�ץX��EXECL
get-eventlog security | export-csv -path d:\EventLog\Event\$FormDate.csv -NoTypeInformation -encoding "unicode"

#Function �NEvent�R��
clear-eventlog security