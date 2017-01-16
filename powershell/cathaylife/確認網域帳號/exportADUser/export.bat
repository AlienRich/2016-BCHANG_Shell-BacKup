#Cxldom09
copy "D:\個人資料\02.各類專案\2015-06-11 金檢查核\網域帳號清查作業\exportADUser\cxldom09.bat" "\\ccsvrdom03\d$\exportADUser\cxldom09.bat" 
psexec \\ccsvrdom03 d:\exportADUser\cxldom09.bat
copy "\\ccsvrdom03\d$\exportADUser\cxldom09.txt" "D:\個人資料\02.各類專案\2015-06-11 金檢查核\網域帳號清查作業\exportADUser\cxldom09.txt"
del "\\ccsvrdom03\d$\exportADUser\cxldom09.txt"
del "\\ccsvrdom03\d$\exportADUser\cxldom09.bat"

#Cxldom04
copy "D:\個人資料\02.各類專案\2015-06-11 金檢查核\網域帳號清查作業\exportADUser\cxldom04.bat" "\\cxlsvr44\d$\exportADUser\cxldom04.bat" 
psexec \\cxlsvr44 d:\exportADUser\cxldom04.bat
copy "\\cxlsvr44\d$\exportADUser\cxldom04.txt" "D:\個人資料\02.各類專案\2015-06-11 金檢查核\網域帳號清查作業\exportADUser\cxldom04.txt"
del "\\cxlsvr44\d$\exportADUser\cxldom04.txt"
del "\\cxlsvr44\d$\exportADUser\cxldom04.bat"

#Cxldom01
copy "D:\個人資料\02.各類專案\2015-06-11 金檢查核\網域帳號清查作業\exportADUser\cxldom01.bat" "\\cxlsvr90\d$\exportADUser\cxldom01.bat" 
psexec \\cxlsvr90 d:\exportADUser\cxldom01.bat
copy "\\cxlsvr90\d$\exportADUser\cxldom01.txt" "D:\個人資料\02.各類專案\2015-06-11 金檢查核\網域帳號清查作業\exportADUser\cxldom01.txt"
del "\\cxlsvr90\d$\exportADUser\cxldom01.txt"
del "\\cxlsvr90\d$\exportADUser\cxldom01.bat"

#Cxldom02
copy "D:\個人資料\02.各類專案\2015-06-11 金檢查核\網域帳號清查作業\exportADUser\cxldom02.bat" "\\cxlsvr160b\d$\exportADUser\cxldom02.bat" 
psexec \\cxlsvr160b d:\exportADUser\cxldom02.bat
copy "\\cxlsvr160b\d$\exportADUser\cxldom02.txt" "D:\個人資料\02.各類專案\2015-06-11 金檢查核\網域帳號清查作業\exportADUser\cxldom02.txt"
del "\\cxlsvr160b\d$\exportADUser\cxldom02.txt"
del "\\cxlsvr160b\d$\exportADUser\cxldom02.bat"

#Cxldom03
copy "D:\個人資料\02.各類專案\2015-06-11 金檢查核\網域帳號清查作業\exportADUser\cxldom03.bat" "\\cxlsvr212\d$\exportADUser\cxldom03.bat" 
psexec \\cxlsvr212 d:\exportADUser\cxldom03.bat
copy "\\cxlsvr212\d$\exportADUser\cxldom03.txt" "D:\個人資料\02.各類專案\2015-06-11 金檢查核\網域帳號清查作業\exportADUser\cxldom03.txt"
del "\\cxlsvr212\d$\exportADUser\cxldom03.txt"
del "\\cxlsvr212\d$\exportADUser\cxldom03.bat"

#Cxlsna
copy "D:\個人資料\02.各類專案\2015-06-11 金檢查核\網域帳號清查作業\exportADUser\cxlsna.bat" "\\202.154.204.57\d$\exportADUser\cxlsna.bat" 
psexec \\202.154.204.57 d:\exportADUser\cxlsna.bat
copy "\\202.154.204.57\d$\exportADUser\cxlsna.txt" "D:\個人資料\02.各類專案\2015-06-11 金檢查核\網域帳號清查作業\exportADUser\cxlsna.txt"
del "\\202.154.204.57\d$\exportADUser\cxlsna.txt"
del "\\202.154.204.57\d$\exportADUser\cxlsna.bat"

#Cxldom00
csvde -d "OU=otheruser,DC=cxldom00,DC=cathay,DC=ins" -r objectClass=user -s cxlsvr104 -u -l displayName,sAMAccountName,userAccountControl,description,scriptPath -o DN -f cxldom00.txt
