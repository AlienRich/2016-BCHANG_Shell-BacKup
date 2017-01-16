
Dim objShell, objNetwork, objSysInfo ,strdepartment, strUser, objUser, strCn

Set objShell = CreateObject("Wscript.Shell")
Set objNetwork = CreateObject("Wscript.Network")
Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
strdepartment = objUser.Get("department")
strCn = objUser.Get("cn")

Select Case LCase(strdepartment)

	Case "台北分公司"
	On Error Resume Next
	objNetwork.MapNetworkDrive "X:", "\\tp0w2k002\" & strdepartment &  "$"
	objNetwork.MapNetworkDrive "h:", "\\tp0w2k002\" & strCn & "" 
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "X:", "\\tp0w2k002\" & strdepartment &  "$"
		objNetwork.MapNetworkDrive "h:", "\\tp0w2k002\" & stCn & ""
	End If

	Case "新竹分公司"
	On Error Resume Next
	objNetwork.MapNetworkDrive "X:", "\\hc0w2k002\" & strdepartment &  "$"
	objNetwork.MapNetworkDrive "h:", "\\hc0w2k002\" & strCn & "" 
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "X:", "\\hc0w2k002\" & strdepartment &  "$"
		objNetwork.MapNetworkDrive "h:", "\\hc0w2k002\" & stCn & ""
	End If

	Case "台中分公司"
	On Error Resume Next
	objNetwork.MapNetworkDrive "X:", "\\tc0w2k002\" & strdepartment &  "$"
	objNetwork.MapNetworkDrive "h:", "\\tc0w2k002\" & strCn & "" 
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "X:", "\\tc0w2k002\" & strdepartment &  "$"
		objNetwork.MapNetworkDrive "h:", "\\tc0w2k002\" & stCn & ""
	End If

	Case "嘉義分公司"
	On Error Resume Next
	objNetwork.MapNetworkDrive "X:", "\\cy0w2k002\" & strdepartment &  "$"
	objNetwork.MapNetworkDrive "h:", "\\cy0w2k002\" & strCn & "" 
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "X:", "\\cy0w2k002\" & strdepartment &  "$"
		objNetwork.MapNetworkDrive "h:", "\\cy0w2k002\" & stCn & ""
	End If	

	Case "台南分公司"
	On Error Resume Next
	objNetwork.MapNetworkDrive "X:", "\\tn0w2k002\" & strdepartment &  "$"
	objNetwork.MapNetworkDrive "h:", "\\tn0w2k002\" & strCn & "" 
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "X:", "\\tn0w2k002\" & strdepartment &  "$"
		objNetwork.MapNetworkDrive "h:", "\\tn0w2k002\" & stCn & ""
	End If

	Case "高雄分公司"
	On Error Resume Next
	objNetwork.MapNetworkDrive "X:", "\\kh0w2k002\" & strdepartment &  "$"
	objNetwork.MapNetworkDrive "h:", "\\kh0w2k002\" & strCn & "" 
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "X:", "\\kh0w2k002\" & strdepartment &  "$"
		objNetwork.MapNetworkDrive "h:", "\\kh0w2k002\" & stCn & ""
	End If

	Case "東部分公司"
	On Error Resume Next
	objNetwork.MapNetworkDrive "X:", "\\hl0w2k002\" & strdepartment &  "$"
	objNetwork.MapNetworkDrive "h:", "\\hl0w2k002\" & strCn & "" 
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "X:", "\\hl0w2k002\" & strdepartment &  "$"
		objNetwork.MapNetworkDrive "h:", "\\hl0w2k002\" & stCn & ""
	End If

	Case else
	On Error Resume Next
	objNetwork.MapNetworkDrive "X:", "\\hq0w23053\" & strdepartment &  "$"
	objNetwork.MapNetworkDrive "h:", "\\hq0w23053\" & strCn & "" 
	If (Err.Number <> 0) Then
		On Error GoTo 0
		objNetwork.RemoveNetworkDrive "i:", True, True
		objNetwork.MapNetworkDrive "X:", "\\hq0w23053\" & strdepartment &  "$"
		objNetwork.MapNetworkDrive "h:", "\\hq0w23053\" & stCn & ""
	End If

End Select

