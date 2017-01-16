
Option Explicit

Dim objNetwork, objShell, objEnv, WSHShell ,objFile

Set objNetwork = CreateObject("Wscript.Network")
Set objShell = CreateObject("Wscript.Shell")
Set objEnv = objShell.Environment("process")
Set WSHShell=Wscript.CreateObject("Wscript.Shell")
Set objFile = CreateObject("Scripting.FileSystemObject")

On Error Resume Next
objFile.DeleteFile("d:\1.txt")



