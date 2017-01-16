' OSType.VBS
Option Explicit
Dim objWMI, objItem, colItems , wsh
Dim strComputer, VerOS, VerBig, Ver9x, Version9x, OS, OSystem

strComputer = "."
Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMI.ExecQuery("Select * from Win32_OperatingSystem",,48)

For Each objItem in colItems
VerBig = Left(objItem.Version,3)
Next

Set wsh = WScript.CreateObject ("WScript.Shell")

Select Case VerBig
Case "5.0"
  wsh.run "I:\SmartIT\ITSetup.exe"
End Select
'Select Case VerBig
'Case "6.0" OSystem = "Vista"
'Case "5.2" OSystem = "Windows 2003"
'Case "5.1" OSystem = "XP"
'Case "5.0" OSystem = "W2K"
'Case "4.0" OSystem = "NT 4.0**"
'Case Else OSystem = "Unknown - probably Win 9x"
'End Select

'Wscript.Echo  VerBig

'Wscript.Echo "Version No : " & VerBig & vbCr _
'& "OS System : " & OSystem

'WScript.Quit

' End of script




