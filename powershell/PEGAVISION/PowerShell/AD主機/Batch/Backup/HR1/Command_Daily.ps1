# 匯入Server 清單
$computers = get-Content C:\Batch\Backup\HR1\HR1.txt

# 將Server清單匯入
  foreach ($computer in $computers) {
  invoke-command -filepath C:\Batch\Backup\HR1\HR1_Daily.Ps1 -computerName $computer
  }