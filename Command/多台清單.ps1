function Get-SystemInfo 
{ 
  param($ComputerName = $env:ComputerName) 
  
      $header = 'Hostname','OSName','OSVersion','OSManufacturer','OSConfig','Buildtype', 'RegisteredOwner','RegisteredOrganization','ProductID','InstallDate', 'StartTime','Manufacturer','Model','Type','Processor','BIOSVersion', 'WindowsFolder' ,'SystemFolder','StartDevice','Culture', 'UICulture', 'TimeZone','PhysicalMemory', 'AvailablePhysicalMemory' , 'MaxVirtualMemory', 'AvailableVirtualMemory','UsedVirtualMemory','PagingFile','Domain' ,'LogonServer','Hotfix','NetworkAdapter' 
      systeminfo.exe /FO CSV /S $ComputerName |  
            Select-Object -Skip 1 |  
            ConvertFrom-CSV -Header $header 
} 

# 匯入Server 清單
$computers = get-Content D:\Command\computer-list.txt

# 將Server清單匯入
  foreach ($computer in $computers) 
      {
        $Info = Get-SystemInfo $computer
        
        $A = New-Object PSObject -Property @{
            Hostname = $Info.Hostname
            OSName = $Info.OSName
            #OSVersion = $Info.OSVersion
            #OSManufacturer = $Info.OSManufacturer
            #OSConfig = $Info.OSConfig
            #Buildtype = $Info.Buildtype
            #RegisteredOwner = $Info.RegisteredOwner
            #RegisteredOrganization = $Info.RegisteredOrganization
            #ProductID = $Info.ProductID
            #InstallDate = $Info.InstallDate
            #StartTime = $Info.StartTime
            Manufacturer = $Info.Manufacturer
            Model = $Info.Model 
            #Type = $Info.Type
            Processor = $Info.Processor
            #BIOSVersion = $Info.BIOSVersion 
            #WindowsFolder = $Info.WindowsFolder
            #SystemFolder = $Info.SystemFolder
            #StartDevice = $Info.StartDevice
            #Culture = $Info.Culture
            #UICulture = $Info.UICulture
            #TimeZone = $Info.TimeZone
            PhysicalMemory = $Info.PhysicalMemory
            #AvailablePhysicalMemory  = $Info.AvailablePhysicalMemory
            #MaxVirtualMemory = $Info.MaxVirtualMemory
            #AvailableVirtualMemory = $Info.AvailableVirtualMemory
            #UsedVirtualMemory = $Info.UsedVirtualMemory
            #PagingFile = $Info.PagingFile
            #Domain = $Info.Domain
            #LogonServer = $Info.LogonServer
            #Hotfix = $Info.Hotfix
            #NetworkAdapter  = $Info.NetworkAdapter
            }
       $a |　select Hostname,OSName,Processor,PhysicalMemory,Manufacturer,Model | export-csv -delimiter "`t" -path D:\Command\Server.csv -Encoding UTF8
       }
     