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

	Case "194" & "99"  'Server'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\10.194.99.242\public$"

	Case "194" & "98"  'User'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\10.194.99.242\public$"
	
	Case else  '¨ä¥L'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\10.194.99.242\public$"
	
	
End Select

