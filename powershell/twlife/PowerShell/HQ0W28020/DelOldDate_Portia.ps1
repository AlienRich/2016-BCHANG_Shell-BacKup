#Function Del old data
Remove-Item -Path D:\其他資料夾\交易作業管理部-PortiaReport\OldData\*

#Function 變數
$FormDate=get-date -Format yyyyMMdd
$SelectFile=Get-ChildItem -Path "D:\其他資料夾\交易作業管理部-PortiaReport" -Recurse | where {$_.CreationTime -lt (get-date).adddays(-100)}  | Where {!$_.PSIsContainer}

#Function Move檔案
$SelectFile | Move-Item -destination D:\其他資料夾\交易作業管理部-PortiaReport\OldData\ 

