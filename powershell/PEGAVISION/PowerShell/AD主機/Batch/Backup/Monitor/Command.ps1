# 匯入Server 清單
$computers = get-Content C:\Batch\Backup\Monitor\Monitor.txt

# 將Server清單匯入
  foreach ($computer in $computers) {
  invoke-command -filepath C:\Batch\Backup\Monitor\Monitor.Ps1 -computerName $computer
  }