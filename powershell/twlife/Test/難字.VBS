
Option Explicit

Dim objNetwork, objShell, objEnv, strServer, WSHShell

Set objNetwork = CreateObject("Wscript.Network")
Set objShell = CreateObject("Wscript.Shell")
Set objEnv = objShell.Environment("process")
Set WSHShell=Wscript.CreateObject("Wscript.Shell")
strServer = objEnv("LOGONSERVER")

WSHShell.RegDelete "HKLM\Software\test"




