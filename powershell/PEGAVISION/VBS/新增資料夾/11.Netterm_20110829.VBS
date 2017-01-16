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

  	Case "1" & "27"  'だそ'
	FileSystemObject.CopyFile "\\tp0w2k002\public\*.ini", WshSysEnv("windir") , True 
  	
	Case "1" & "10"  '穝λだそ''λ璢跋场'
	FileSystemObject.CopyFile "\\hc0w2k002\public\*.ini", WshSysEnv("windir") , True 
	
 	Case "1" & "11"  'いだそ'
 	FileSystemObject.CopyFile "\\tc0w2k002\public\*.ini", WshSysEnv("windir") , True 
	
	Case "1" & "12"  '堕跋场'
	FileSystemObject.CopyFile "\\ty0w23001\public\*.ini", WshSysEnv("windir") , True 

	Case "1" & "13"  '古竡だそ'
	FileSystemObject.CopyFile "\\cy0w2k002\public\*.ini", WshSysEnv("windir") , True 

	Case "1" & "14"  '玭だそ'
	FileSystemObject.CopyFile "\\tn0w2k002\public\*.ini", WshSysEnv("windir") , True 

	Case "1" & "15"  '蔼动だそ'
	FileSystemObject.CopyFile "\\ks0w2k002\public\*.ini", WshSysEnv("windir") , True 
	
	Case "1" & "16" '狥场だそ''狥场跋场'
	FileSystemObject.CopyFile "\\hl0w2k003\public\*.ini", WshSysEnv("windir") , True 

	Case "16" & "62" 'い材跋场'
	FileSystemObject.CopyFile "\\tc1w23001\public\*.ini", WshSysEnv("windir") , True 
	objNetwork.MapNetworkDrive "i:", "\\tc1w23001\public"

	Case "1" & "20"  '关穨厩'
	FileSystemObject.CopyFile "\\hq3w23001\public\*.ini", WshSysEnv("windir") , True 
	
	Case "1" & "26" '狥跋场'
	FileSystemObject.CopyFile "\\pt0w23001\public\*.ini", WshSysEnv("windir") , True 

	Case "1" & "22" '狝加'
	FileSystemObject.CopyFile "\\hq2w28002\ImportAP$\*.ini", WshSysEnv("windir") , True 

	Case "1" & "23" '狝加'
	FileSystemObject.CopyFile "\\hq2w28002\ImportAP$\*.ini", WshSysEnv("windir") , True 

	Case "1" & "24" '狝加'
	FileSystemObject.CopyFile "\\hq2w28002\ImportAP$\*.ini", WshSysEnv("windir") , True 
	
	Case "1" & "28"  '刮︽現场''刮穨叭场''蝗︽玂繧场'
	FileSystemObject.CopyFile "\\hq1w23001\public\*.ini", WshSysEnv("windir") , True 
	
	Case else  'ㄤ'
	FileSystemObject.CopyFile "\\hq0w28020\ImportAP$\*.ini", WshSysEnv("windir") , True 

	
End Select

