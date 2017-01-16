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

	
	Case "1" & "12"  '桃園區部'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\ty0w23001\public"

	Case "16" & "62" '台中第一區部'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tc1w23001\public"

	Case "1" & "20"  '台壽企業大學'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hq3w23001\public"
	
	Case "1" & "26" '屏東區部'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\pt0w23001\public"

	Case else  '其他'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\twlife\fileserver\Public"
	
	
End Select

