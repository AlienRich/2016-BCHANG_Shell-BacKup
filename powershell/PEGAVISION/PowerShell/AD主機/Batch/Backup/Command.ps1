# 匯入Server 清單
$computers = get-Content C:\Batch\Backup\Windows2008.txt

# 將Server清單匯入
  foreach ($computer in $computers) {
  #invoke-command -filepath C:\Batch\Backup\test.Ps1 -computerName $computer
  invoke-command -filepath C:\Batch\Backup\Windows2008.Ps1 -computerName $computer
  }