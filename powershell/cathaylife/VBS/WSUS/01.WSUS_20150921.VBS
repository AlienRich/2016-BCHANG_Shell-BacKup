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


Select Case LCase(IP1) & (IP2)	'�P�_��X'

	Case "10" & "80" '�x�_��F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"

	Case "10" & "81" '�x�_��F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"

	Case "10" & "82" '�x�_��F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"

	Case "10" & "83" '�x�_��F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"

	Case "10" & "84" '�x�_��F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"

	Case "10" & "85" '�x�_��F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"

	Case "10" & "86" '�x�_��F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"

	Case "10" & "89" '�x�_��F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"

	Case "10" & "90" '�x�_��F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"	

	Case "10" & "3" '��˦�F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr453"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr453"

	Case "10" & "88" '��˦�F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr453"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr453"

	Case "10" & "4" '�x����F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr454"

	Case "10" & "96" '�x����F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr454"

	Case "10" & "97" '�x����F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr454"

	Case "10" & "98" '�x����F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr454"

	Case "10" & "99" '�x����F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr454"

	Case "10" & "100" '�x����F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr454"

	Case "10" & "101" '�x����F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr454"

	Case "10" & "102" '�x����F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr454"

	Case "10" & "103" '�x����F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr454"

	Case "10" & "104" '�x����F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr454"

	Case "10" & "4" '�x����F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr454"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr454"

	Case "10" & "5" '�x�n��F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr456"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr456"

	Case "10" & "112" '�x�n��F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr456"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr456"

	Case "10" & "114" '�x�n��F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr456"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr456"

	Case "10" & "7" '������F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr457"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr457"

	Case "10" & "120" '������F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr457"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr457"

	Case "10" & "121" '������F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr457"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr457"

	Case "10" & "122" '������F����'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr457"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr457"	


	Case else         '�P�_�T�X'
		Select Case LCase(IP1) & (IP2) & (IP3)

  		Case "10" & "87" & "32" '�`���q'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"

  		Case "10" & "87" & "33" '�`���q'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"

  		Case "10" & "87" & "34" '�`���q'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"

  		Case "10" & "87" & "35" '�`���q'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"

  		Case "10" & "87" & "36" '�`���q'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"

  		Case "10" & "87" & "37" '�`���q'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"

  		Case "10" & "87" & "38" '�`���q'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr451"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr451"

  		Case "10" & "2" & "0" '�x�_��F����'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"

  		Case "10" & "2" & "1" '�x�_��F����'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"

  		Case "10" & "87" & "15" '�Q�s����'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"

  		Case "10" & "87" & "17" '�H���Ш|����'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"

  		Case "10" & "87" & "21" '�x�_�V�m�B'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr452"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr452"

  		Case "10" & "87" & "18" '�������@'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr453"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr453"

  		Case "10" & "87" & "19" '�������G'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr453"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr453"

  		Case "10" & "180" & "104" '��A����'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://10.180.103.216"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://10.180.103.216"

  		Case "10" & "180" & "105" '��A����'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://10.180.103.216"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://10.180.103.216"

  		Case "10" & "180" & "106" '��A����'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://10.180.103.216"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://10.180.103.216"

  		Case "10" & "180" & "107" '��A����'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://10.180.103.216"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://10.180.103.216"

		Case else         '��L'
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://cxlsvr450"
		WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://cxlsvr450"
		End Select

		
End Select

WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer",00000001,"REG_DWORD"

Set WSHShell=nothing