' WSUS
' --------------------------------------------------------' 
Option Explicit
Dim objNetwork, WSHShell , objWMIService , colItems  ,objitem 
Dim strIPAddress , IPAddress , strcomputer , IP1 , IP2 , IP3, IP4
' ---------------------------------------------------------'
Set WSHShell = CreateObject("Wscript.Shell")
strcomputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

for each objitem in colitems

strIPAddress = Join(objitem.IPAddress, ",")
IPAddress = Split(strIPAddress,".")
IP1 =  IPAddress(0)
IP2 =  IPAddress(1)
IP3 =  IPAddress(2)
IP4 =  IPAddress(3)
Next

Select Case LCase(IP2) & (IP3)

  	Case "1" & "10" '新竹分公司'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://HC0W28003"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://HC0W28003"

  	Case "1" & "11" '台中分公司'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://TC0W28003"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://TC0W28003"

	Case "1" & "13"  '嘉義分公司'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://CY0W28003"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://CY0W28003"
	
	Case "1" & "14"   '台南分公司'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://TN0W28003"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://TN0W28003"

	Case "1" & "15"   '高雄分公司'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://KH0W28003"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://KH0W28003"

	Case "1" & "16"    '東部分公司'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://HL0W28003"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://HL0W28003"
	
	Case "1" & "22" '客服'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://HQ2W28003"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://HQ2W28003"

	Case "1" & "23" '客服'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://HQ2W28003"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://HQ2W28003"

	Case "1" & "24" '直效'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://HQ2W28003"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://HQ2W28003"

	Case "1" & "27" '台北分公司'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://TP0W28003"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://TP0W28003"

	Case "1" & "28" '第二事業中心'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://TP0W28003"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://TP0W28003"

	Case "1" & "20" '台壽企業大'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://HQ3W28003"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://HQ3W28003"

	Case "16" & "62" '台中中一區部'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://TC1W28003"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://TC1W28003"

	Case "1" & "12" '桃園區部'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://TY0W28003"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://TY0W28003"

	Case else         '其他'
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\ElevateNonAdmins",00000001,"REG_DWORD"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUServer","http://HQ0W28007"
	WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\WUStatusServer","http://HQ0W28007"

End Select

WSHShell.RegWrite "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer",00000001,"REG_DWORD"

Set WSHShell=nothing
