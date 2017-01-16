' VB Script Document
Option Explicit
Dim objShell, objNetwork ,objFSO
Set objNetwork = WScript.CreateObject("WScript.Network") 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
set objShell = WScript.CreateObject("Wscript.Shell") 

On Error Resume Next 
objNetwork.RemoveNetworkDrive "H:",True,True

' End of Script
