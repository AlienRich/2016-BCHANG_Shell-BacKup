' MapPublic.VBS
' --------------------------------------------------------' 
Option Explicit
Dim objNetwork, objShell, objWMIService , colItems  ,objitem 
Dim strIPAddress , IPAddress , strcomputer , IP1 , IP2 , IP3, IP4
Dim FileSystemObject , WshSysEnv , WshShell
' ---------------------------------------------------------'

Set objNetwork = CreateObject("Wscript.Network")
Set objShell = CreateObject("Wscript.Shell")
strcomputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
set WshShell = WScript.CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("Process") 

for each objitem in colitems

strIPAddress = Join(objitem.IPAddress, ",")
IPAddress = Split(strIPAddress,".")
IP1 =  IPAddress(0)
IP2 =  IPAddress(1)
IP3 =  IPAddress(2)
IP4 =  IPAddress(3)
Next

Select Case LCase(IP2) & (IP3)

  	Case "1" & "27"  '台北分公司'
	FileSystemObject.CopyFile "\\tp0w28002\ImportAP$\NetTerm\*.ini", WshSysEnv("windir") , True
	FileSystemObject.CopyFile "\\tp0w28002\ImportAP$\NetTerm\*.exe", "c:\NetTerm\" , True 
  	
	Case "1" & "10"  '新竹分公司''竹苗區部'
	FileSystemObject.CopyFile "\\hc0w28002\ImportAP$\NetTerm\*.ini", WshSysEnv("windir") , True 
	FileSystemObject.CopyFile "\\hc0w28002\ImportAP$\NetTerm\*.exe", "c:\NetTerm\" , True 
  	
 	Case "1" & "11"  '台中分公司'
 	FileSystemObject.CopyFile "\\tc0w28002\ImportAP$\NetTerm\*.ini", WshSysEnv("windir") , True 
	FileSystemObject.CopyFile "\\tc0w28002\ImportAP$\NetTerm\*.exe", "c:\NetTerm\" , True 
  	
	Case "1" & "12"  '桃園區部'
	FileSystemObject.CopyFile "\\ty0w28002\ImportAP$\NetTerm\*.ini", WshSysEnv("windir") , True 
	FileSystemObject.CopyFile "\\ty0w28002\ImportAP$\NetTerm\*.exe", "c:\NetTerm\" , True 
  	
	Case "1" & "13"  '嘉義分公司'
	FileSystemObject.CopyFile "\\cy0w28002\ImportAP$\NetTerm\*.ini", WshSysEnv("windir") , True 
	FileSystemObject.CopyFile "\\cy0w28002\ImportAP$\NetTerm\*.exe", "c:\NetTerm\" , True 
  	
	Case "1" & "14"  '台南分公司'
	FileSystemObject.CopyFile "\\tn0w28002\ImportAP$\NetTerm\*.ini", WshSysEnv("windir") , True 
	FileSystemObject.CopyFile "\\tn0w28002\ImportAP$\NetTerm\*.exe", "c:\NetTerm\" , True 
  	
	Case "1" & "15"  '高雄分公司'
	FileSystemObject.CopyFile "\\kh0w28002\ImportAP$\NetTerm\*.ini", WshSysEnv("windir") , True 
	FileSystemObject.CopyFile "\\kh0w28002\ImportAP$\NetTerm\*.exe", "c:\NetTerm\" , True 
  	
	Case "1" & "16" '東部分公司''東部區部'
	FileSystemObject.CopyFile "\\hl0w28002\ImportAP$\NetTerm\*.ini", WshSysEnv("windir") , True 
	FileSystemObject.CopyFile "\\hl0w28002\ImportAP$\NetTerm\*.exe", "c:\NetTerm\" , True 
  	
	Case "16" & "62" '台中第一區部'
	FileSystemObject.CopyFile "\\tc1w28002\ImportAP$\NetTerm\*.ini", WshSysEnv("windir") , True 
	FileSystemObject.CopyFile "\\tc1w28002\ImportAP$\NetTerm\*.exe", "c:\NetTerm\" , True 
  	
	Case "1" & "20"  '台壽企業大學'
	FileSystemObject.CopyFile "\\hq3w28002\ImportAP$\NetTerm\*.ini", WshSysEnv("windir") , True 
	FileSystemObject.CopyFile "\\hq3w28002\ImportAP$\NetTerm\*.exe", "c:\NetTerm\" , True 
  	
	Case else  '其他'
	FileSystemObject.CopyFile "\\hq0w28020\ImportAP$\NetTerm\*.ini", WshSysEnv("windir") , True 
	FileSystemObject.CopyFile "\\hq0w28020\ImportAP$\NetTerm\*.exe", "c:\NetTerm\" , True 
  	
	
End Select

