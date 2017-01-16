' Disable_Windows10
' --------------------------------------------------------' 
Option Explicit
Dim WSHShell
' ---------------------------------------------------------'
Set WSHShell = CreateObject("Wscript.Shell")

WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\DisableOSUpgrade",00000001,"REG_DWORD"


