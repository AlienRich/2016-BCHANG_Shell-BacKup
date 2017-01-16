' MapPublic.VBS
' --------------------------------------------------------' 
Option Explicit
Dim filesys
' ---------------------------------------------------------'


set filesys=CreateObject("Scripting.FileSystemObject") 

If not filesys.FileExists("C:\subinacl.exe") Then 
	filesys.CopyFile "\\cxldom00\netlogon\subinacl.exe", "C:\" 
End If 



	


