set Exe_name=*.rpt
set DesPath=\c$\Program Files\FeibApp\Report\
set Desfilename=r0094.rpt
Set RptSourcePath="%~dp0%Exe_name%"
Set SysDate=%date:~0,4%%date:~5,2%%date:~8,2%

echo on
for %%i in (%Desfilename%) do rename "\\%1%DesPath%\%%i" %%i.%SysDate%
copy %RptSourcePath% "\\%1%DesPath%"
echo off
pause