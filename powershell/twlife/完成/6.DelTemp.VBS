' DelFiles.vbs
' --------------------------------------------------------' 
Option Explicit
Dim objFSO, objFiles, objShell, intCount
Dim strFile, strName, strDirectory, strEnv, strExt
' ---------------------------------------------------------'

Set objShell = CreateObject("Wscript.Shell")
'strEnv = "%USERPROFILE%\Local Settings\Temp"
strEnv = objShell.ExpandEnvironmentStrings("%temp%")
intcount = 0

' Create the File System Object
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO = objFSO.GetFolder(strEnv)
Set objFiles = objFSO.Files
For each strFile in objFiles
On Error Resume Next
    strFile.delete 
    intCount = intCount +1
next

set objFSO = nothing
set strFile = nothing
set objFiles = nothing
'WScript.Echo intCount & " files were deleted"
WScript.Quit
