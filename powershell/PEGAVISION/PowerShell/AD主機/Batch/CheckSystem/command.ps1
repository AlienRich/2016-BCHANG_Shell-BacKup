# 匯入Server 清單
$computers = get-Content C:\Batch\CheckSystem\computer-list.txt
$Time = get-date -forma HH
$Date = get-date -format yyyyMMdd

#確認資料夾，如沒有就產生
if (!(Test-Path -Path "\\10.194.99.242\PublicFolder\管理部\資訊部\資訊共用\test\$date\")) {  
  # 建立目錄  
  New-Item -Path "\\10.194.99.242\PublicFolder\管理部\資訊部\資訊共用\test\$date\"  -ItemType "directory" 
}
# Lets create our variables
$HostInfo = @()

# 將Server清單匯入
  foreach ($computer in $computers) {
     # 抓取Server Cpu使用情況
     $ProcessorStats = Get-WmiObject win32_processor -computer $computer | select -exp LoadPercentage
     # 多核心的平均值
     $ComputerCpu = ($ProcessorStats|measure -Average).Average
     # Lets create a re-usable WMI method for memory stats
     $OperatingSystem = Get-WmiObject win32_OperatingSystem -computer $computer
     # Lets grab the free memory
     $FreeMemory = $OperatingSystem.FreePhysicalMemory
     # Lets grab the total memory
     $TotalMemory = $OperatingSystem.TotalVisibleMemorySize
     # Lets do some math for percent
     $MemoryUsed = ($TotalMemory - $FreeMemory) / $TotalMemory * 100
     $PercentMemoryUsed = "{0:N2}" -f $MemoryUsed
     #抓取硬碟資訊
     $Disk=Get-WmiObject win32_logicaldisk -Filter "DriveType = 3" -computer $computer | select-object DeviceID, VolumeName,Size,FreeSpace,@{Name="UseSpace";Expression={100-$_.FreeSpace/$_.Size*100}}
     #顯示硬碟使用
     $UseDisk = $Disk | select DeviceID
     $UseSpace = $Disk | select UseSpace

     # Lets throw them into an object for outputting
     $ComputerCpu=[string]$ComputerCpu+"%"
     $PercentMemoryUsed=[string]$PercentMemoryUsed+"%"
     $objHostInfo = New-Object System.Object
     $objHostInfo | Add-Member -MemberType NoteProperty -Name Name -Value $computer
     $objHostInfo | Add-Member -MemberType NoteProperty -Name CPU% -Value $ComputerCpu
     $objHostInfo | Add-Member -MemberType NoteProperty -Name MemoryUse% -Value $PercentMemoryUsed
     $a="";     
     if($UseDisk -is [array]){	
     	for ($i=0;$i -lt $UseDisk.length; $i++) {                 
        	 $a+=$UseDisk[$i]
		 $b=$UseSpace[$i]
         	 $b=$b -replace "@{UseSpace=", ""
	   	 $b=$b -replace "}", ""
	         $a=$a -replace "@{DeviceID=", ""
		 $a=$a -replace "}", ""
		 
	         $a+=[int]$b
		 $a+="% "
	     }
     }else{
	$a+=$UseDisk
        $b=$UseSpace
	$b=$b -replace "@{UseSpace=", ""
	$b=$b -replace "}", ""
        $a=$a -replace "@{DeviceID=", ""
	$a=$a -replace "}", ""
        $a+=[int]$b
        $a+="% "
     }
     $objHostInfo | Add-Member -MemberType NoteProperty -Name DiskUse% -Value $a
     # Lets dump our info into an array
     $HostInfo+= $objHostInfo
    
}


# Lets output to the console
$HostInfo | out-file C:\Batch\CheckSystem\Check.txt

Copy-Item "C:\Batch\CheckSystem\Check.txt" "\\10.194.99.242\PublicFolder\管理部\資訊部\資訊共用\test\$date\$Time.txt"
#$ArrayA = Get-Content D:\Batch\CheckSystem\Date.txt
#$ArrayB =  Get-Content D:\Batch\CheckSystem\Date1.txt
#[array]$ArrayC = $ArrayA[1,2,3,4,5,6,7] + $ArrayB[3,4,5,6] + $ArrayA[8..100]
#Out-File  -InputObject  $ArrayC  -FilePath D:\Batch\CheckSystem\Check.txt
#Function Send-Mail
Send-MailMessage -smtpserver PV-EX-CASHUB02 -to billchang@pegavision.com -from CheckSystem@pegavision.com -Subject "CheckSystem" -Encoding ([System.Text.Encoding]::UTF8) -Attachment "C:\Batch\CheckSystem\Check.txt"

Remove-Item C:\Batch\CheckSystem\Check.txt

