' --------------------------------------------------------' 
Option Explicit
Dim objNetwork, objShell, objSysInfo, objUser
Dim strServer, strdepartment, strUser,strPassword
' --------------------------------------------------------' 
strPassword="654321"
Set objShell = CreateObject("Wscript.Shell")
Set objNetwork = CreateObject("Wscript.Network")
Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)


strPassword = InputBox ("Enter Password")
 
 
 
On Error Resume Next
 
 
 
Set objUser = GetObject("WinNT://" & strUser ,User) 

objUser.SetPassword strPassword 

objUser.SetInfo 

 
 
If Err.Number <> 0 Then
 
MsgBox "Error: An Incorrect Domain Or User Name Was Entered!"
 
Err.Clear
 
Else
 
MsgBox "The Password Has Been Changed For " & UCase(strDomain) & "\" & UCase(strUser) 

End If
