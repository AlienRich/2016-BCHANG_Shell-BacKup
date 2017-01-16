#Function ≈‹º∆
$FormDate=get-date -Format yyyyMMdd
$SelectFile=Get-ChildItem -Path D:\BackupForld\

cd "c:\Program Files (x86)\7-Zip\"
.\7z.exe a -tzip D:\BackupForld\$FormDate.zip D:\BackupForld\

move D:\BackupForld\$FormDate.zip "\\hq0w23065\LogBackup\EventLog"

Remove-Item D:\BackupForld\* -recurse 