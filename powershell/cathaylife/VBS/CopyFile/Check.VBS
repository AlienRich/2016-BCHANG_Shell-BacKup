' MapPublic.VBS
' --------------------------------------------------------' 
Option Explicit
Dim filesys
' ---------------------------------------------------------'


set filesys=CreateObject("Scripting.FileSystemObject") 

If not filesys.FileExists("C:\subinacl.exe") Then 
	Wscript.Echo "�L�ɮ�" 

Else
    Wscript.Echo "���ɮ�"
End If



	


