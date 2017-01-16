
Option Explicit

Dim objNetwork, objShell, objEnv, strServer, WSHShell

Set objNetwork = CreateObject("Wscript.Network")
Set objShell = CreateObject("Wscript.Shell")
Set objEnv = objShell.Environment("process")
Set WSHShell=Wscript.CreateObject("Wscript.Shell")
strServer = objEnv("LOGONSERVER")

Select Case LCase(strServer)
	Case "\\hq0w23027"
	WSHShell.RegWrite "HKCU\TEST\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKCU\TEST\WUServer","http://10.1.2.92"
	WSHShell.RegWrite "HKCU\TEST\WUStatusServer","http://10.1.2.92"
	
	Case "\\hq0w23028"
	WSHShell.RegWrite "HKCU\TEST\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKCU\TEST\WUServer","http://10.1.2.92"
	WSHShell.RegWrite "HKCU\TEST\WUStatusServer","http://10.1.2.92"

	Case "tp0w23001"
	WSHShell.RegWrite "HKCU\TEST\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKCU\TEST\WUServer","http://10.1.27.23"
	WSHShell.RegWrite "HKCU\TEST\WUStatusServer","http://10.1.27.23"

	Case "\\hc0w23001"
	WSHShell.RegWrite "HKCU\TEST\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKCU\TEST\WUServer","http://10.1.10.23"
	WSHShell.RegWrite "HKCU\TEST\WUStatusServer","http://10.1.10.23"

	Case "\\tc0w23001"
	WSHShell.RegWrite "HKCU\TEST\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKCU\TEST\WUServer","http://10.1.11.23"
	WSHShell.RegWrite "HKCU\TEST\WUStatusServer","http://10.1.11.23"

	Case "\\cy0w23001"
	WSHShell.RegWrite "HKCU\TEST\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKCU\TEST\WUServer","http://10.1.13.23"
	WSHShell.RegWrite "HKCU\TEST\WUStatusServer","http://10.1.13.23"
	
	Case "\\tn0w23001"
	WSHShell.RegWrite "HKCU\TEST\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKCU\TEST\WUServer","http://10.1.14.23"
	WSHShell.RegWrite "HKCU\TEST\WUStatusServer","http://10.1.14.23"

	Case "\\kh0w23001"
	WSHShell.RegWrite "HKCU\TEST\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKCU\TEST\WUServer","http://10.1.15.23"
	WSHShell.RegWrite "HKCU\TEST\WUStatusServer","http://10.1.15.23"

	Case "\\hl0w23001"
	WSHShell.RegWrite "HKCU\TEST\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKCU\TEST\WUServer","http://10.1.16.23"
	WSHShell.RegWrite "HKCU\TEST\WUStatusServer","http://10.1.16.23"
End Select

WSHShell.RegWrite "HKCU\TEST\AU\NoAUShutdownOption",00000001, "REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\NoAUAsDefaultShutdownOption",00000000,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\NoAutoUpdate",00000000,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\UseWUServer",00000001,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\RescheduleWaitTimeEnabled",00000001,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\NoAutoRebootWithLoggedOnUsers",00000001,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\DetectionFrequencyEnabled",00000001,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\AutoInstallMinorUpdates",00000001,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\RebootWarningTimeoutEnabled",00000001,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\RebootRelaunchTimeoutEnabled",00000001,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\RescheduleWaitTime",00000005,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\DetectionFrequency",00000004,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\RebootWarningTimeout",00000005,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\AUOptions",00000004,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\ScheduledInstallDay",00000000,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\RebootRelaunchTimeout",000000010,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\AU\ScheduledInstallTime",00000012,"REG_DWORD"
