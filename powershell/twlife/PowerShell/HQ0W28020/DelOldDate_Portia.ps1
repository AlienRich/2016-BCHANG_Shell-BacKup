#Function Del old data
Remove-Item -Path D:\��L��Ƨ�\����@�~�޲z��-PortiaReport\OldData\*

#Function �ܼ�
$FormDate=get-date -Format yyyyMMdd
$SelectFile=Get-ChildItem -Path "D:\��L��Ƨ�\����@�~�޲z��-PortiaReport" -Recurse | where {$_.CreationTime -lt (get-date).adddays(-100)}  | Where {!$_.PSIsContainer}

#Function Move�ɮ�
$SelectFile | Move-Item -destination D:\��L��Ƨ�\����@�~�޲z��-PortiaReport\OldData\ 

