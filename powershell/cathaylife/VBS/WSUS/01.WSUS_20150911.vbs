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

Select Case LCase(IP1) & (IP2) & (IP3)

  	Case "10" & "87" & "32" '�`���q'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR451-�`���q"

	Case "10" & "87" & "33" '�`���q'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR451-�`���q"
	
	Case "10" & "87" & "34" '�`���q'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR451-�`���q"

	Case "10" & "87" & "35" '�`���q'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR451-�`���q"
	
	Case "10" & "87" & "36" '�`���q'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR451-�`���q"

	Case "10" & "87" & "37" '�`���q'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR451-�`���q"

	Case "10" & "87" & "38" '�`���q'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR451-�`���q"

	Case "10" & "87" & "39" '�`���q'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR451-�`���q"

	Case "10" & "2" & "0" '�_�@��'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR452-�x�_��F����"

	Case "10" & "2" & "1" '�_�@��'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR452-�x�_��F����"
	
	Case "10" & "3" & "0"  '��˦�'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr453"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr453"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR452-��˦�F����"

	Case "10" & "3" & "1"  '��˦�'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr453"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr453"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR452-��˦�F����"
	
	Case "10" & "4" & "0"   '�x����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR452-�x����F����"

	Case "10" & "4" & "1"   '�x����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR452-�x����F����"

	Case "10" & "6" & "0"   '�x�n��'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr456"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr456"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR452-�x�n��F����"

	Case "10" & "6" & "1"   '�x�n��'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr456"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr456"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR452-�x�n��F����"

	Case "10" & "7" & "20"    '������'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr457"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr457"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR452-�x������F����"

	Case "10" & "7" & "21"    '������'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr457"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr457"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR452-�x������F����"
	
	Case else         '��L'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr450"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr450"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\TargetGroup","CXLSVR450-����"

End Select

WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer",00000001,"REG_DWORD"

Set WSHShell=nothing