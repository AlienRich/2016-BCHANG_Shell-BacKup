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
strdepartment = objUser.Get("department")
strCn = objUser.Get("sAMAccountName")

On Error Resume Next 
objNetwork.RemoveNetworkDrive "X:",True,True
On Error Resume Next 
objNetwork.RemoveNetworkDrive "H:",True,True

Select Case LCase(strdepartment)

	Case  "董事會"
	On Error Resume Next
	objNetwork.MapNetworkDrive "X:", "\\hq0w23053\董事會$"
	On Error Resume Next 
	objNetwork.MapNetworkDrive "h:", "\\hq0w23053\" & strCn & ""

End Select



