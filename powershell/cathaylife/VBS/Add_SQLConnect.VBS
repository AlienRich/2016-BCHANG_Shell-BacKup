' Disable_Windows10
' --------------------------------------------------------' 
Option Explicit
Dim WSHShell
' ---------------------------------------------------------'
Set WSHShell = CreateObject("Wscript.Shell")

WSHShell.RegWrite "HKLM\SOFTWARE\Microsoft\MSSQLServer\Client\ConnectTo\cxlsvr22","DBMSSOCN,cxlsvr22-new,1433","REG_SZ"
WSHShell.RegWrite "HKLM\SOFTWARE\Microsoft\MSSQLServer\Client\ConnectTo\cxlsvr55","DBMSSOCN,cxlsvr23-new,1433","REG_SZ"
WSHShell.RegWrite "HKLM\SOFTWARE\Microsoft\MSSQLServer\Client\ConnectTo\cxlsvr80","DBMSSOCN,cxlsvr22-new,1433","REG_SZ"
WSHShell.RegWrite "HKLM\SOFTWARE\Microsoft\MSSQLServer\Client\ConnectTo\cxlsvr60","DBMSSOCN,cxlsvr219,1433","REG_SZ"
WSHShell.RegWrite "HKLM\SOFTWARE\Wow6432Node\Microsoft\MSSQLServer\Client\ConnectTo\cxlsvr22","DBMSSOCN,cxlsvr22-new,1433","REG_SZ"
WSHShell.RegWrite "HKLM\SOFTWARE\Wow6432Node\Microsoft\MSSQLServer\Client\ConnectTo\cxlsvr55","DBMSSOCN,cxlsvr23-new,1433","REG_SZ"
WSHShell.RegWrite "HKLM\SOFTWARE\Wow6432Node\Microsoft\MSSQLServer\Client\ConnectTo\cxlsvr80","DBMSSOCN,cxlsvr22-new,1433","REG_SZ"
WSHShell.RegWrite "HKLM\SOFTWARE\Wow6432Node\Microsoft\MSSQLServer\Client\ConnectTo\cxlsvr60","DBMSSOCN,cxlsvr219,1433","REG_SZ"


