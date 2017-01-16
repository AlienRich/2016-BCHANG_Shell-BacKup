' MapPublic.VBS
' --------------------------------------------------------' 
Option Explicit
Dim objNetwork, objShell, objWMIService , colItems  ,objitem 
Dim strIPAddress , IPAddress , strcomputer , IP1 , IP2 , IP3, IP4
' ---------------------------------------------------------'

Set objNetwork = CreateObject("Wscript.Network")
Set objShell = CreateObject("Wscript.Shell")
strcomputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

for each objitem in colitems

strIPAddress = Join(objitem.IPAddress, ",")
IPAddress = Split(strIPAddress,".")
IP1 =  IPAddress(0)
IP2 =  IPAddress(1)
IP3 =  IPAddress(2)
IP4 =  IPAddress(3)
Next

On Error Resume Next 
objNetwork.RemoveNetworkDrive "i:",True,True

Select Case LCase(IP2) & (IP3)

  	Case "1" & "27"  'だそ'
	On Error Resume Next 
	objNetwork.RemoveNetworkDrive "j:",True,True
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tp0w2k002\public"

  	Case "1" & "10"  '穝λだそ''λ璢跋场'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hc0w2k002\public"
	
 	Case "1" & "11"  'いだそ'
  	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tc0w2k002\public"
	
	Case "1" & "12"  '堕跋场'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\ty0w23001\public"

	Case "1" & "13"  '古竡だそ'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\cy0w2k002\public"

	Case "1" & "14"  '玭だそ'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tn0w2k002\public"

	Case "1" & "15"  '蔼动だそ'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\ks0w2k002\public"
	
	Case "1" & "16" '狥场だそ''狥场跋场'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hl0w2k003\public"

	Case "16" & "62" 'い材跋场'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tc1w23001\public"

	Case "1" & "20"  '关穨厩'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hq3w23001\public"
	
	Case "1" & "26" '狥跋场'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\pt0w23001\public"

	Case "1" & "22" '狝加'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\twlife\fileserver\Public"

	Case "1" & "23" '狝加'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\twlife\fileserver\Public"

	Case "1" & "24" '狝加'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\twlife\fileserver\Public"
	
	Case "1" & "28"  '刮︽現场''刮穨叭场''蝗︽玂繧场'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hq1w23001\public"
	
	Case else  'ㄤ'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\twlife\fileserver\Public"
	
	
End Select

