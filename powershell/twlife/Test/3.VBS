Option Explicit

Dim objSysInfo, strSite

Set objSysInfo = CreateObject("ADSystemInfo")
strSite = objSysInfo.SiteName
Select Case LCase(strSite)
Case "mysite1"
' .... code to map drive and printer.
Case "mysite2"
' .... code to map drive and printer.
End Select