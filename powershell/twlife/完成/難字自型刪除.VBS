
Option Explicit

Dim objNetwork, objShell, objEnv, WSHShell ,objFile

Set objNetwork = CreateObject("Wscript.Network")
Set objShell = CreateObject("Wscript.Shell")
Set objEnv = objShell.Environment("process")
Set WSHShell=Wscript.CreateObject("Wscript.Shell")
Set objFile = CreateObject("Scripting.FileSystemObject")

On Error Resume Next
WSHShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts\New Gulim (TrueType)"
On Error Resume Next
objFile.DeleteFile("c:\windows\Fonts\NGULIM.TTF")



