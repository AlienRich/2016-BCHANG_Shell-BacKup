
Option Explicit

Dim objNetwork, objShell, objEnv, WSHShell ,objFile

Set objNetwork = CreateObject("Wscript.Network")
Set objShell = CreateObject("Wscript.Shell")
Set objEnv = objShell.Environment("process")
Set WSHShell=Wscript.CreateObject("Wscript.Shell")
Set objFile = CreateObject("Scripting.FileSystemObject")

On Error Resume Next
objFile.DeleteFile("C:\NetTerm\WINET.EXE")
objFile.DeleteFile("C:\WINDOWS\netterm.ini")
FileSystemObject.CopyFile "I:\Apps\netterm\netterm.ini", "C:\WINDOWS\"
FileSystemObject.CopyFile "I:\Apps\netterm\pnpg.exe", "C:\WINDOWS\system32\"
FileSystemObject.CopyFile "I:\Apps\netterm\wp2.exe", "C:\WINDOWS\system32\"


