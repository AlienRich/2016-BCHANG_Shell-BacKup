Option Explicit

Dim objNetwork, objShell, objEnv, strServer, WSHShell

Set objNetwork = CreateObject("Wscript.Network")
Set objShell = CreateObject("Wscript.Shell")
Set objEnv = objShell.Environment("process")
Set WSHShell=Wscript.CreateObject("Wscript.Shell")
strServer = objEnv("LOGONSERVER")

WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable",00000001,"REG_DWORD"
WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyOverride","10.*.*.*,172.*.*.*,192.168.*.*,<local>,","REG_SZ"
WSHShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer","proxy.local:80","REG_SZ"