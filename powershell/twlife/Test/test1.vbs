' WSUS.VBS
' --------------------------------------------------------' 
Option Explicit
Dim objNetwork, objEnv, WSHShell,objShell
Dim strServer
' ---------------------------------------------------------'
set WshShell = WScript.CreateObject("WScript.Shell")  
Set objNetwork = CreateObject("Wscript.Network")
Set objShell = CreateObject("Wscript.Shell")
Set objEnv = objShell.Environment("process")
strServer = objEnv("LOGONSERVER")


WScript.Echo strServer

