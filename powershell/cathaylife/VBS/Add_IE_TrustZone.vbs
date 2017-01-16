Dim WSHShell
Set WSHShell = WScript.CreateObject("WScript.Shell")

'增加SharePoint IE安全信信任
WSHShell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2\1406", 0,"REG_DWORD"

'增加CXLSVR1167 IE安全信信任
WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\cxlsvr1167.cxldom00.cathay.ins\http", 2,"REG_DWORD"
Set WSHShell = Nothing