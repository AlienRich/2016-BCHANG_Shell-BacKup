' MapPublic.VBS + Desktop
' --------------------------------------------------------' 
Option Explicit
Dim objNetwork, objShell, objEnv
Dim strServer
' ---------------------------------------------------------'

Set objNetwork = CreateObject("Wscript.Network")
Set objShell = CreateObject("Wscript.Shell")
Set objEnv = objShell.Environment("process")
strServer = objEnv("LOGONSERVER")

On Error Resume Next 
objNetwork.RemoveNetworkDrive "i:",True,True

Select Case LCase(strServer)
	Case "\\hq0w23027","\\hq0w23028"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hq0w23053\public"
	
  Case "tp0w23001"
  On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tp0w2k002\public"

	Case "\\hc0w23001"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hc0w2k002\public"

	Case "\\tc0w23001"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tc0w2k002\public"

	Case "\\cy0w23001"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\cy0w2k002\public"
	
	Case "\\tn0w23001"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tn0w2k002\public"

	Case "\\kh0w23001"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\kh0w2k002\public"

	Case "\\hl0w23001"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hl0w2k003\public"
End Select

