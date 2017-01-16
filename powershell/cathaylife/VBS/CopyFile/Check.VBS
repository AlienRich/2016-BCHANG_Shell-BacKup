' MapPublic.VBS
' --------------------------------------------------------' 
Option Explicit
Dim filesys
' ---------------------------------------------------------'


set filesys=CreateObject("Scripting.FileSystemObject") 

If not filesys.FileExists("C:\subinacl.exe") Then 
	Wscript.Echo "無檔案" 

Else
    Wscript.Echo "有檔案"
End If



	


