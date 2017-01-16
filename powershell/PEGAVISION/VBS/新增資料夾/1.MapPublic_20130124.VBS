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

	Case "1" & "10"  '新竹分公司'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hc0w28002\public$"
	
	Case "1" & "11"  '台中分公司'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tc0w28002\public$"
	
	Case "1" & "12"  '桃園區部'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\ty0w28002\public"

	Case "1" & "13"  '嘉義分公司'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\cy0w28002\public$"

	Case "1" & "14"  '台南分公司'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tn0w28002\public$"

	Case "1" & "15"  '高雄分公司'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\kh0w28002\public$"

	Case "1" & "16"  '東部分公司'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hl0w28002\public$"

	Case "16" & "62" '台中第一區部'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tc1w28002\public"
	
	Case "1" & "27" '台北分公司'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tp0w28002\public"

	Case "1" & "20"  '台壽企業大學'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hq3w28002\public$"

	Case else  '其他'
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hq0w28020\public$"
	
	
End Select

