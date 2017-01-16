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
strCn = objUser.Get("cn")

On Error Resume Next 
objNetwork.RemoveNetworkDrive "W:",True,True

Select Case LCase(strdepartment)

	Case  "╰参砏购场","莱ノ╰参场","╰参キ场","计狝叭场"
	On Error Resume Next
	objNetwork.MapNetworkDrive "W:", "\\hq0w23053\筿福いみ$"

End Select



