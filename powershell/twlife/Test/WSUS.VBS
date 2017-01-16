Option Explicit

Dim WSHShell 
Set WSHShell=Wscript.CreateObject("Wscript.Shell")

WSHShell.RegWrite "HKCU\TEST\ElevateNonAdmins",00000001,"REG_DWORD"
WSHShell.RegWrite "HKCU\TEST\WUServer","http://10.1.2.92"
WSHShell.RegWrite "HKCU\TEST\WUStatusServer","http://10.1.2.92"
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
Set WSHShell=nothing