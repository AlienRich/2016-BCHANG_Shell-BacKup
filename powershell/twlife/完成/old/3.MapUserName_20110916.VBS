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
objNetwork.RemoveNetworkDrive "h:",True,True

Select Case LCase(strdepartment)

	Case  "だそ","跋场","ㄆ穨矪","癡絤场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\tp0w2k002\" & strCn &  "$"
	
	Case "穝λだそ","λ璢跋场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\hc0w2k002\" & strCn &  "$"
  
  	Case "堕跋场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\ty0w23001\" & strCn &  "$"

	Case "いだそ","い材跋场","い场ㄆ穨矪","い癡絤场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\tc0w2k002\" & strCn &  "$"
  
  	Case "い材跋场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\tc1w23001\" & strCn &  "$"

	Case "古竡だそ","冻古跋场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\cy0w2k002\" & strCn &  "$"

	Case "玭だそ","玭跋场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\tn0w2k002\" & strCn &  "$"


	Case "蔼动だそ","蔼动跋场","玭场ㄆ穨矪","蔼动癡絤场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\ks0w2k002\" & strCn &  "$"
  
  	Case "狥跋场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\pt0w23001\" & strCn &  "$" 

	Case "狥场だそ","狥场跋场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\hl0w2k003\" & strCn &  "$"
  
	Case  "穨叭珇借场","︽綪场","玂禣场","め狝叭场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\hq2w28002\" & strCn &  "$"

	Case  "毙▅癡絤场"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\hq3w23001\" & strCn &  "$"
	
	Case else
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\hq0w28020\" & strCn &  "$"


End Select



