$Dir = "c:\batch"
$copy = "C:\batch\robocopy.exe"
$Format = Get-Date -Format yyyyMMdd

& $copy Z:\ D:\ /R:1 /W:0 /s >> C:\Batch\$Format.txt  