' WSUS
' --------------------------------------------------------' 
Option Explicit
Dim objNetwork, WSHShell , objWMIService , colItems  ,objitem 
Dim strIPAddress , IPAddress , strcomputer , IP1 , IP2 , IP3, IP4
' ---------------------------------------------------------'
Set WSHShell = CreateObject("Wscript.Shell")
strcomputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

for each objitem in colitems

strIPAddress = Join(objitem.IPAddress, ",")
IPAddress = Split(strIPAddress,".")
IP1 =  IPAddress(0)
IP2 =  IPAddress(1)
IP3 =  IPAddress(2)
IP4 =  IPAddress(3)
Next

Select Case LCase(IP2) & (IP3)

  	Case "1" & "100" 
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://hq0w28007"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://hq0w28007"

	Case else         '¥þ³¡'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://hq0w28007"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://hq0w28007"

End Select

WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer",00000001,"REG_DWORD"

Set WSHShell=nothing
