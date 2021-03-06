' MapDep&UserName.vbs
' --------------------------------------------------------' 
Option Explicit
Dim objNetwork, objShell, objSysInfo, objUser
Dim strServer, strdepartment, strUser, strCn
' --------------------------------------------------------' 

Set objShell = CreateObject("Wscript.Shell")
Set objNetwork = CreateObject("Wscript.Network")
Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
On Error Resume Next 
strdepartment = objUser.Get("department")
strCn = objUser.Get("sAMAccountName")

On Error Resume Next 
objNetwork.RemoveNetworkDrive "X:",True,True
On Error Resume Next 
objNetwork.RemoveNetworkDrive "H:",True,True

Select Case LCase(strdepartment)

	Case  "だそ","跋场","ㄆ穨矪","癡絤场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\tp0w2k002\" & strdepartment &  "$"
	
	Case "穝λだそ","λ璢跋场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\hc0w2k002\" & strdepartment &  "$"
  
  	Case "堕跋场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\ty0w23001\" & strdepartment &  "$"

	Case "いだそ","い材跋场","い场ㄆ穨矪","い癡絤场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\tc0w2k002\" & strdepartment &  "$"
  
  	Case "い材跋场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\tc1w23001\" & strdepartment &  "$"

	Case "古竡だそ","冻古跋场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\cy0w2k002\" & strdepartment &  "$"

	Case "玭だそ","玭跋场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\tn0w2k002\" & strdepartment &  "$"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\tn0w2k002\" & strCn & "" 

	Case "蔼动だそ","蔼动跋场","玭场ㄆ穨矪","蔼动癡絤场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\ks0w2k002\" & strdepartment &  "$"
  
  	Case "狥跋场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\pt0w23001\" & strdepartment &  "$" 

	Case "狥场だそ","狥场跋场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\hl0w2k003\浆だそ$"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\hl0w2k003\" & strCn & ""
  
 	Case  "赋ㄆ穦"
	On Error Resume Next
	objNetwork.MapNetworkDrive "X:", "\\hq0w23053\赋ㄆ穦$"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\hq0w23053\" & strCn & ""

  	Case  "计瞶场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\hq0w23053\计瞶场$"
  	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\hq0w23022\坝珇祇いみ$"
	
	Case  "计瞶场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\hq0w23022\计瞶场$"
  	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\hq0w23022\坝珇祇いみ$"
	
	Case  "╰参砏购场","莱ノ╰参场","╰参キ场","计狝叭场"
	On Error Resume Next
	objNetwork.MapNetworkDrive "W:", "\\hq0w23053\筿福いみ$" 
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\hq0w23053\" & strdepartment &  "$"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\hq0w23053\" & strCn & ""

	Case  "穨叭珇借场","︽綪场","玂禣场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\hq2w28002\" & strdepartment &  "$"

	Case  "め狝叭场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\hq2w28002\" & strdepartment &  "$"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "W:", "\\hq2w28002\Service$"
	
	Case  "毙▅癡絤场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\hq3w23001\" & strdepartment &  "$"
	
	Case  "刮玂篬︽現场","刮玂篬穨叭场","蝗︽玂繧场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\hq1w23001\" & strdepartment &  "$"

	Case else
	On Error Resume Next 
	objNetwork.MapNetworkDrive "X:", "\\hq0w23053\" & strdepartment &  "$"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\hq0w23053\" & strCn & "" 

End Select



