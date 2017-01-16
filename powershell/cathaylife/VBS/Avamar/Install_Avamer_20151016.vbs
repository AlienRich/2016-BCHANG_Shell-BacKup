'判斷執行檔，如果有avscc.exe既不執行
set service = GetObject ("winmgmts:")

for each Process in Service.InstancesOf ("Win32_Process")
	If Process.Name = "avscc.exe" then
	Wscript.Quit	
	end if	
Next

'如果沒有avscc.exe開始動作
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

'增加資料夾
strDirectory = "D:\個人電腦自動備份"
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

'判斷OS版本
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

'判斷IP	
Select Case LCase(IP1) & (IP2)	'判斷兩碼'
	Case "10" & "80" '台北行政中心'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "81" '台北行政中心'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "82" '台北行政中心'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "83" '台北行政中心'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "84" '台北行政中心'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "85" '台北行政中心'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "86" '台北行政中心'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "89" '台北行政中心'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "90" '台北行政中心'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "126" '台北行政中心'
	Filesys.CopyFile TPPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"		

	Case "10" & "3" '桃竹行政中心'
	Filesys.CopyFile TYPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "88" '桃竹行政中心'
	Filesys.CopyFile TYPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "4" '台中行政中心'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "96" '台中行政中心'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "97" '台中行政中心'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "98" '台中行政中心'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "99" '台中行政中心'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "100" '台中行政中心'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "101" '台中行政中心'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "102" '台中行政中心'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "103" '台中行政中心'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "104" '台中行政中心'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "4" '台中行政中心'
	Filesys.CopyFile TCPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "6" '台南行政中心'
	Filesys.CopyFile TNPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "112" '台南行政中心'
	Filesys.CopyFile TNPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "114" '台南行政中心'
	Filesys.CopyFile TNPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "128" '台南行政中心'
	Filesys.CopyFile TNPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "7" '高雄行政中心'
	Filesys.CopyFile KHPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "120" '高雄行政中心'
	Filesys.CopyFile KHPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "121" '高雄行政中心'
	Filesys.CopyFile KHPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"

	Case "10" & "122" '高雄行政中心'
	Filesys.CopyFile KHPAth, strSysDrive & "\"		
	WSHShell.run strSysDrive & "\setup.exe"	

	Case else         '判斷三碼'
		Select Case LCase(IP1) & (IP2) & (IP3)

		Case "10" & "87" & "32" '總公司'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "33" '總公司'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "34" '總公司'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "35" '總公司'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "36" '總公司'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "37" '總公司'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "38" '總公司'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "2" & "0" '台北行政中心'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "2" & "1" '台北行政中心'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "26" '台北行政中心'
		Filesys.CopyFile HQPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "15" '松山機場'
		Filesys.CopyFile TPPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "17" '淡水教育中心'
		Filesys.CopyFile TPPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "21" '台北訓練處'
		Filesys.CopyFile TPPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "18" '桃園機場一'
		Filesys.CopyFile TYPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "87" & "19" '桃園機場二'
		Filesys.CopyFile TYPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

		Case "10" & "180" & "103" '資服中心'
		Filesys.CopyFile ITPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "180" & "104" '資服中心'
		Filesys.CopyFile ITPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "180" & "105" '資服中心'
		Filesys.CopyFile ITPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "180" & "106" '資服中心'
		Filesys.CopyFile ITPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

  		Case "10" & "180" & "107" '資服中心'
		Filesys.CopyFile ITPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"

		Case else         '其他'
		Filesys.CopyFile SEPPAth, strSysDrive & "\"		
		WSHShell.run strSysDrive & "\setup.exe"
		End Select
	End Select	
Wscript.Quit		
