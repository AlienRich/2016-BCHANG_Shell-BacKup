Dim WSHShell
Set WSHShell = WScript.CreateObject("WScript.Shell")
WSHShell.RegWrite "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\SecureProtocols", 2688,"REG_DWORD"
Set WSHShell = Nothing