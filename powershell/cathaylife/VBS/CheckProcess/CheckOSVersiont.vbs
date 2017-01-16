Set WshShell = CreateObject("WScript.Shell")
 
OsType = WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
 
If OsType = "x86" then
wscript.echo "Windows 32bit system detected"
elseif OsType = "AMD64" then
wscript.echo "Windows 64bit system detected"
end if