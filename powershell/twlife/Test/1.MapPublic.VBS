
Option Explicit

Dim objNetwork, objShell, objEnv, strServer, objSysInfo ,strdepartment, strUser, objUser, strCn


Set objNetwork = CreateObject("Wscript.Network")

Set objShell = CreateObject("Wscript.Shell")
Set objEnv = objShell.Environment("process")
strServer = objEnv("LOGONSERVER")
Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
strdepartment = objUser.Get("department")
strCn = objUser.Get("cn")

Select Case LCase(strServer)
	Case "\\hq0w23027"
	if ( " strdepartment " = " 台中第一區部 " ) then
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tc1w23001\public"
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "i:", "\\tc1w23001\public"	
	else
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hq0w23053\public"
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "i:", "\\hq0w23053\public"
	End If
	End If
	
	Case "\\hq0w23028"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hq0w23053\public"
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "i:", "\\hq0w23053\public"
	End If

	Case "tp0w23001"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tp0w2k002\public"
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "i:", "\\tp0w2k002\public"
	End If

	Case "\\hc0w23001"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hc0w2k002\public"
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "i:", "\\hc0w2k002\public"
	End If

	Case "\\tc0w23001"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tc0w2k002\public"
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "i:", "\\tc0w2k002\public"
	End If

	Case "\\cy0w23001"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\cy0w2k002\public"
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "i:", "\\cy0w2k002\public"
	End If
	
	Case "\\tn0w23001"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\tn0w2k002\public"
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "i:", "\\tn0w2k002\public"
	End If

	Case "\\kh0w23001"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\kh0w2k002\public"
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "i:", "\\kh0w2k002\public"
	End If

	Case "\\hl0w23001"
	On Error Resume Next
	objNetwork.MapNetworkDrive "i:", "\\hl0w2k003\public"
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "i:", "\\hl0w2k003\public"
	End If
End Select