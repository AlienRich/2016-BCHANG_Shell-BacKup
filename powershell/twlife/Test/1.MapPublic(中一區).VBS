
Option Explicit

Dim objNetwork, objShell, objEnv, strServer

Set objNetwork = CreateObject("Wscript.Network")

Set objShell = CreateObject("Wscript.Shell")
Set objEnv = objShell.Environment("process")
strServer = objEnv("LOGONSERVER")

Select Case LCase(strServer)
	Case "\\tc0w23001"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tc1w23001\public"
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "i:", "\\tc1w23001\public"
	End If

End Select