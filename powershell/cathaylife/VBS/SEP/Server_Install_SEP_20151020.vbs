'�P�_�����ɡA�p�G��Smc.exe�J������
set service = GetObject ("winmgmts:")

for each Process in Service.InstancesOf ("Win32_Process")
	If Process.Name = "Smc.exe" then
	Wscript.Quit	
	end if	
Next

'�p�G�S��Smc.exe�}�l�ʧ@
Dim objNetwork, WSHShell , objWMIService , colItems  ,objitem , path  
Dim strIPAddress , IPAddress , strcomputer , IP1 , IP2 , IP3, IP4
Dim HQPath,TPPath,TYPath,TCPath,TNPath,KHPath,ITPAth,SEPPath,strEXE
Dim strSysDrive

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

'�P�_OS����
OsType = WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
 
If OsType = "x86" then
HQPath	= "\\cxlsvr471\sep11\setup.exe"
TPPath	= "\\cxlsvr472\sep11\setup.exe"
TYPath	= "\\cxlsvr473\sep11\setup.exe"
TCPath	= "\\cxlsvr474\sep11\setup.exe"
TNPath	= "\\cxlsvr476\sep11\setup.exe"
KHPath	= "\\cxlsvr477\sep11\setup.exe"
ITPath	= "\\sep11\sep11\setup.exe"
SEPPath	= "\\cxlsvr169\package\setup.exe"
strEXE = "setup.exe"
elseif OsType = "AMD64" then
HQPath	= "\\cxlsvr471\sep11\setup64.exe"
TPPath	= "\\cxlsvr472\sep11\setup64.exe"
TYPath	= "\\cxlsvr473\sep11\setup64.exe"
TCPath	= "\\cxlsvr474\sep11\setup64.exe"
TNPath	= "\\cxlsvr476\sep11\setup64.exe"
KHPath	= "\\cxlsvr477\sep11\setup64.exe"
ITPath	= "\\sep11\sep11\setup64.exe"
SEPPath	= "\\cxlsvr169\package\setup64.exe"
strEXE = "setup64.exe"
end if

set filesys=CreateObject("Scripting.FileSystemObject") 
strSysDrive = WshShell.ExpandEnvironmentStrings("%systemdrive%")

'�P�_IP	
Select Case LCase(IP1) & (IP2)	'�P�_��X'
	
	Case "10" & "122" '������F����'
	Filesys.CopyFile SEPPath, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\" & strEXE

	Case else         '�P�_�T�X'
		Select Case LCase(IP1) & (IP2) & (IP3)

		Case "10" & "87" & "32" '�`���q'
		Filesys.CopyFile SEPPath, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\" & strEXE

  		
		Case else         '��L'
		Filesys.CopyFile SEPPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\" & strEXE
		End Select
	End Select	
Wscript.Quit		
