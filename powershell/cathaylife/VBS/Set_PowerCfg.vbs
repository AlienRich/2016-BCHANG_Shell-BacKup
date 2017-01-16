Dim wshShell
Set wshShell = CreateObject("WScript.Shell")
wshShell.Exec("powercfg -change -monitor-timeout-ac 10")
Set wshShell = Nothing

