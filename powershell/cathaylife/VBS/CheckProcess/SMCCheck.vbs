set service = GetObject ("winmgmts:")

for each Process in Service.InstancesOf ("Win32_Process")
	If Process.Name = "Sm.exe" then
	wscript.echo "running"
	Wscript.Quit	
	end if	
Next

Dim objNetwork, WSHShell , objWMIService , colItems  ,objitem 
Dim strIPAddress , IPAddress , strcomputer , IP1 , IP2 , IP3, IP4

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
	wscript.echo strIPAddress &" is 1 running on computer "
	
	Case "10" & "81" '�x�_��F����'
	wscript.echo strIPAddress &" is 2 running on computer "

	Case else         '�P�_�T�X'
		Select Case LCase(IP1) & (IP2) & (IP3)

		Case "10" & "180" & "104" '�`���q'
		wscript.echo strIPAddress & " is 3 running on computer "

		End Select
	End Select	
Wscript.Quit		
