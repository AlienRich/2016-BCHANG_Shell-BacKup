# 匯入Server 清單
$computers = get-Content C:\Batch\Backup\PegaCloud\PegaCloud.txt

# 將Server清單匯入
  foreach ($computer in $computers) {
  #invoke-command -filepath C:\Batch\Backup\test.Ps1 -computerName $computer
  invoke-command -filepath C:\Batch\Backup\PegaCloud\PegaCloud.Ps1 -computerName $computer
  }