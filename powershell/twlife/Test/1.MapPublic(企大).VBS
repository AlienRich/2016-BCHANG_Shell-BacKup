
Option Explicit

Dim objNetwork, objShell, objEnv, strServer

Set objNetwork = CreateObject("Wscript.Network")

Set objShell = CreateObject("Wscript.Shell")
Set objEnv = objShell.Environment("process")
strServer = objEnv("LOGONSERVER")

Select Case LCase(strServer)
	Case "\\hq0w23027"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hq3w23001\public"
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "i:", "\\hq0w23053\public"
	End If
	
	Case "\\hq0w23028"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hq3w23001\public"
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "i:", "\\hq0w23053\public"
	End If
End Select