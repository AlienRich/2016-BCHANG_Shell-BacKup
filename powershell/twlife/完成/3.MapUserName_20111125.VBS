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


	Case  "系統平台部"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\hq0w28020\" & strCn &  "$"
	
	Case else
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\hq0w28020\" & strCn &  "$"


End Select



