# 匯入Server 清單
$computers = get-Content .\computer-list1.txt

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
$HostInfo | out-file d:\Date1.txt
#Function Send-Mail
#Send-MailMessage -smtpserver hq0w23074 -to bchang@twlife.com.tw,jerrywu@twlife.com.tw -from SIM@twlife.com.tw -Subject "CheckList" -Encoding ([System.Text.Encoding]::UTF8) -Attachment "date.txt"
#Remove-Item d:\Date.txt

