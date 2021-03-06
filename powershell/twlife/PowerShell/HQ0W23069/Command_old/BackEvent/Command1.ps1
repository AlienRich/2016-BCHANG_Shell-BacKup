# 匯入Server 清單
$computers = get-Content D:\Command\computer-list1.txt

# Lets create our variables
# $HostInfo = @()

# 將Server清單匯
  Foreach ($computer in $computers) {
    $EventLogs = get-Content D:\Command\BackEvent\EventLog.txt
    $BackupFolder = "D:\BackupForld\" + (Get-Date).tostring("yyyyMMdd")+ "\"  + $computer
    New-Item -ItemType directory -Path $BackupFolder  
       IF ( $computer -eq "HQ0W23042")
            { $DiskDrive = "N" }
       ELSE
            { $DiskDrive = "C" }    
        
    Foreach ( $logname in $EventLogs ) {

       $BackupFile = $DiskDrive + ":\"  + $logname +".evt"
       $logFile = Get-WmiObject Win32_NTEventlogFile -computer $computer | Where-Object {$_.logfilename -eq $logname}
       $logFile.PSBase.Scope.Options.EnablePrivileges = $true
       $logFile.backupeventlog($BackupFile)
       $S = "\\" + $computer + "\" + $DiskDrive + "$\" + $logname + ".evt"
       $d = "D:\BackupForld\" + (Get-Date).tostring("yyyyMMdd")+ "\"  + $computer + "\"
       move $S $d
    }

}

