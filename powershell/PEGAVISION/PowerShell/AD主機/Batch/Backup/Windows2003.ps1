# 匯入Server 清單
$computers = get-Content C:\Batch\Backup\Windows2003.txt

# 將Server清單匯入
  foreach ($computer in $computers) {
  psexec \\$computer D:\Batch\Backup\1.bat
}
