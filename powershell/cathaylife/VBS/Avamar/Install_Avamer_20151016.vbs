'�P�_�����ɡA�p�G��avscc.exe�J������
set service = GetObject ("winmgmts:")

for each Process in Service.InstancesOf ("Win32_Process")
	If Process.Name = "avscc.exe" then
	Wscript.Quit	
	end if	
Next

'�p�G�S��avscc.exe�}�l�ʧ@
Dim objNetwork, WSHShell , objWMIService , colItems  ,objitem , path  
Dim strIPAddress , IPAddress , strcomputer , IP1 , IP2 , IP3, IP4
Dim HQPath,TPPath,TYPath,TCPath,TNPath,KHPath,ITPAth,SEPPath
Dim strSysDrive
Dim objFSO, objFolder, objShell, strDirectory

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

'�W�[��Ƨ�
strDirectory = "D:\�ӤH�q���۰ʳƥ�"
Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FolderExists(strDirectory) Then
   Set objFolder = objFSO.GetFolder(strDirectory)
   WScript.Echo strDirectory & " already Exists "
Else
   Set objFolder=objFSO.CreateFolder(strDirectory)
   WScript.Echo "Just created " & strDirectory
End If

If err.number = vbEmpty then
   Set objShell = CreateObject("WScript.Shell")
   'objShell.run ("Explorer" &" " & strDirectory & "\" )
Else 
   WScript.echo "VBScript Error: " & err.number
End If

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
SEPPath	= "\\cxlsvr203\sep11\setup.exe"
elseif OsType = "AMD64" then
HQPath	= "\\cxlsvr471\sep11\setup64.exe"
TPPath	= "\\cxlsvr472\sep11\setup64.exe"
TYPath	= "\\cxlsvr473\sep11\setup64.exe"
TCPath	= "\\cxlsvr474\sep11\setup64.exe"
TNPath	= "\\cxlsvr476\sep11\setup64.exe"
KHPath	= "\\cxlsvr477\sep11\setup64.exe"
ITPath	= "\\sep11\sep11\setup64.exe"
SEPPath	= "\\cxlsvr203\sep11\setup64.exe"
end if

set filesys=CreateObject("Scripting.FileSystemObject") 
strSysDrive = WshShell.ExpandEnvironmentStrings("%systemdrive%")

'�P�_IP	
Select Case LCase(IP1) & (IP2)	'�P�_��X'
	Case "10" & "80" '�x�_��F����'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "81" '�x�_��F����'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "82" '�x�_��F����'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "83" '�x�_��F����'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "84" '�x�_��F����'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "85" '�x�_��F����'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "86" '�x�_��F����'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "89" '�x�_��F����'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "90" '�x�_��F����'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "126" '�x�_��F����'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"		

	Case "10" & "3" '��˦�F����'
	Filesys.CopyFile TYPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "88" '��˦�F����'
	Filesys.CopyFile TYPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "4" '�x����F����'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "96" '�x����F����'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "97" '�x����F����'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "98" '�x����F����'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "99" '�x����F����'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "100" '�x����F����'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "101" '�x����F����'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "102" '�x����F����'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "103" '�x����F����'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "104" '�x����F����'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "4" '�x����F����'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "6" '�x�n��F����'
	Filesys.CopyFile TNPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "112" '�x�n��F����'
	Filesys.CopyFile TNPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "114" '�x�n��F����'
	Filesys.CopyFile TNPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "128" '�x�n��F����'
	Filesys.CopyFile TNPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "7" '������F����'
	Filesys.CopyFile KHPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "120" '������F����'
	Filesys.CopyFile KHPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "121" '������F����'
	Filesys.CopyFile KHPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "122" '������F����'
	Filesys.CopyFile KHPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"	

	Case else         '�P�_�T�X'
		Select Case LCase(IP1) & (IP2) & (IP3)

		Case "10" & "87" & "32" '�`���q'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "33" '�`���q'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "34" '�`���q'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "35" '�`���q'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "36" '�`���q'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "37" '�`���q'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "38" '�`���q'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "2" & "0" '�x�_��F����'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "2" & "1" '�x�_��F����'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "26" '�x�_��F����'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "15" '�Q�s����'
		Filesys.CopyFile TPPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "17" '�H���Ш|����'
		Filesys.CopyFile TPPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "21" '�x�_�V�m�B'
		Filesys.CopyFile TPPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "18" '�������@'
		Filesys.CopyFile TYPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "19" '�������G'
		Filesys.CopyFile TYPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

		Case "10" & "180" & "103" '��A����'
		Filesys.CopyFile ITPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "180" & "104" '��A����'
		Filesys.CopyFile ITPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "180" & "105" '��A����'
		Filesys.CopyFile ITPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "180" & "106" '��A����'
		Filesys.CopyFile ITPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "180" & "107" '��A����'
		Filesys.CopyFile ITPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

		Case else         '��L'
		Filesys.CopyFile SEPPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"
		End Select
	End Select	
Wscript.Quit		
