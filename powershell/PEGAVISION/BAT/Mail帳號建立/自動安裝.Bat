"REM No it hasn't"
IF EXIST "D:\Mail\Outlook.pst" GOTO :END

reg delete "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook" /f

md d:\mail
Copy \\10.194.99.242\Public$\Tools\Mail帳號建立\Mail\outlook.pst d:\mail

REM No it hasn't
IF EXIST "C:\Program Files (x86)\Microsoft Office\Office14\outlook.exe" GOTO :x86

"C:\Program Files\Microsoft Office\Office14\outlook.exe" /importprf \\10.194.99.242\public$\Tools\Mail帳號建立\Config\Outlook.PRF
ping 1.1.1.1 -n 1 -w 6000
pskill outlook.exe
regedit /s \\10.194.99.242\public$\Tools\Mail帳號建立\Config\Config.reg
START "Outlook" "C:\Program Files\Microsoft Office\Office14\outlook.exe"

:x86
START "Outlook" "C:\Program Files (x86)\Microsoft Office\Office14\outlook.exe" /importprf \\10.194.99.242\public$\Tools\Mail帳號建立\Config\Outlook.PRF
ping 1.1.1.1 -n 1 -w 6000
pskill outlook.exe
regedit /s \\10.194.99.242\public$\Tools\Mail帳號建立\Config\Config.reg
START "Outlook" "C:\Program Files (x86)\Microsoft Office\Office14\outlook.exe"
EXIT

:END
msg %username% Check D:\Mail\Outlook.pst
EXIT
