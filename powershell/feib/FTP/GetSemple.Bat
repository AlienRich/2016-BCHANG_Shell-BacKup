@echo off

::�w�qFTP�n�J�b���K�X
SET TASKNAME=FPGtoFEIB_YYYYMMDD_FECR

SET server=172.24.254.210
SET username=FeibtoFpg
SET password=D278V239F
SET serverpath=/UCCC/


::�w�q����榡
SET YYYYMMDD=%date:~0,4%%date:~5,2%%date:~8,2%
SET /A Date0=%YYYYMMDD%


::�w�q�ɮ׮榡
SET File1=FECR%Date0%
SET File2=FTCR%Date0%
::SET File2=ACCLINK_APPLY_0999_0805_%Date0%_01.sign
SET JOBNAME1=%File1%
SET FTPLOG=D:\FTPjob\LOG\%File1%_%Date0%.TXT


::�w�qFTP�ʧ@�ɦW
SET FTPActions=%File1%.FTP

::FTP�ʧ@
echo open %server%>>%FTPActions%
echo %username%>>%FTPActions%
echo %password%>>%FTPActions%
echo bin>>%FTPActions%
echo lcd D:\RootFTP\HP\FPGtoFEIB>>%FTPActions%
echo cd %serverpath%>>%FTPActions%
echo get %File1%>>%FTPActions%
echo get %File2%>>%FTPActions%
echo bye>>%FTPActions%


::�}�l����
echo.
echo %Date0% %time%>>%FTPLOG%
echo ========================>>%FTPLOG%
ftp -s:%FTPActions%>>%FTPLOG%
echo ========================>>%FTPLOG%
echo %Date0% %time%>>%FTPLOG%
echo.
del %FTPActions%

