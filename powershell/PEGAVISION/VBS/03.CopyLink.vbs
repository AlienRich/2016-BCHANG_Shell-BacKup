Dim fso, wshShell, strUserProfile

Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

strA = "\\10.194.99.242\tool$\Temp\*.website"
strB = WshShell.ExpandEnvironmentStrings("%USERPROFILE%\")  & "Desktop"
strC = "D:\"

fso.CopyFile strA , strB & "\" , True