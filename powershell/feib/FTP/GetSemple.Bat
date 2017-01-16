@echo off

::定義FTP登入帳號密碼
SET TASKNAME=FPGtoFEIB_YYYYMMDD_FECR

SET server=172.24.254.210
SET username=FeibtoFpg
SET password=D278V239F
SET serverpath=/UCCC/


::定義日期格式
SET YYYYMMDD=%date:~0,4%%date:~5,2%%date:~8,2%
SET /A Date0=%YYYYMMDD%


::定義檔案格式
SET File1=FECR%Date0%
SET File2=FTCR%Date0%
::SET File2=ACCLINK_APPLY_0999_0805_%Date0%_01.sign
SET JOBNAME1=%File1%
SET FTPLOG=D:\FTPjob\LOG\%File1%_%Date0%.TXT


::定義FTP動作檔名
SET FTPActions=%File1%.FTP

::FTP動作
echo open %server%>>%FTPActions%
echo %username%>>%FTPActions%
echo %password%>>%FTPActions%
echo bin>>%FTPActions%
echo lcd D:\RootFTP\HP\FPGtoFEIB>>%FTPActions%
echo cd %serverpath%>>%FTPActions%
echo get %File1%>>%FTPActions%
echo get %File2%>>%FTPActions%
echo bye>>%FTPActions%


::開始執行
echo.
echo %Date0% %time%>>%FTPLOG%
echo ========================>>%FTPLOG%
ftp -s:%FTPActions%>>%FTPLOG%
echo ========================>>%FTPLOG%
echo %Date0% %time%>>%FTPLOG%
echo.
del %FTPActions%


