#Function 變數
$SelectFile=Get-ChildItem -Path "D:\部門資料夾\資訊服務中心-系統平台部\USB資料交換區\" -Recurse | where {$_.CreationTime -lt (get-date).adddays(-3)}  | Where {!$_.PSIsContainer}

#Function Move檔案
$SelectFile | Remove-Item

