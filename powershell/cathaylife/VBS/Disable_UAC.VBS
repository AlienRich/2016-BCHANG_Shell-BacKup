' Disable_UAC
' --------------------------------------------------------' 
Option Explicit
Dim WSHShell
' ---------------------------------------------------------'
Set WSHShell = CreateObject("Wscript.Shell")

WSHShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA",00000000,"REG_DWORD"


