' Author: Frank Hsieh 
' Corp: Chang Cheng Information Consultant Co.,LTD 
' Ver(0.9)
' ==========================================================================
Option Explicit
On Error Resume Next
Dim WshShell  
Dim objFSO
Dim objEnv   
Dim DCServer
Dim sUserName
Dim strTitle,strPrompt,strDefault,NotesPassword 
Dim strCmdMail,StrcmdCalendar 
Dim PSTFile
Dim NotesID,NotesDB,DominoServer
Dim ExchServerName
Dim NamesFile
Dim strComputer
Dim objWMIService
Dim colComputer
Dim objComputer
Dim arr1
Dim ArrUserName
Dim strLogonServer
Dim objDomain
Dim strDC
Dim strADsPath

'Check if outlook has been installed.
CheckOutlookInstlled()
CheckHasNotesClient()
CheckTransendMigrator()

Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Get AD User Name and Logon DC to check if the user's mailbox has been created.
Set objEnv = WshShell.Environment("PROCESS")

'Add "c:\Migrate\tm" to system path 
If InStr(1,objEnv("PATH"),"C:\Migrate\tm") = 0 Then 
objEnv("PATH")=objEnv("PATH") & ";" & "C:\Migrate\tm" 
End If

DCServer = Trim(WshShell.RegRead("HKCU\Volatile Environment\LOGONSERVER"))
DCServer = Right(DCServer,Len(DCServer)-2)

sUserName = objEnv("USERNAME")

Call CheckHasMailbox(sUserName,DCServer)

NamesFile = GetContact()

ExchServerName =GetExchangeServerName(sUserName,DCServer)

Call ContactMigration(NamesFile,ExchServerName,sUserName)

Public Function QueryActiveDirectory(sUserName)
'Function:      QueryActiveDirectory
'Purpose:       Search the Active Directory's Global Catalog for users
'Parameters:    UserName - user to search for
'Return:        The user's distinguished name
 
    Dim oAD 'As IADs
    Dim oGlobalCatalog 'As IADs
    Dim oRecordSet 'As Recordset
    Dim oConnection 'As New Connection
    Dim strADsPath 'As String
    Dim strQuery 'As String
    Dim strUPN 'As String
    set oRecordSet = CreateObject("ADODB.Recordset")
    set oConnection = CreateObject("ADODB.Connection")

    'Determine the global catalog path
    Set oAD = GetObject("GC:")
    For Each oGlobalCatalog In oAD
        strADsPath = oGlobalCatalog.AdsPath
    Next
    'Initialize the ADO object
    oConnection.Provider = "ADsDSOObject"
    'The ADSI OLE-DB provider
    oConnection.Open "ADs Provider"
    'Create the search string
    strQuery = "<" & strADsPath & _
      ">;(&(objectClass=user)(objectCategory=person)(samaccountName=" & _
      sUserName & "));userPrincipalName,cn,distinguishedName;subtree"  
        'Execute the query

    Set oRecordSet = oConnection.Execute(strQuery)
    If oRecordSet.EOF And oRecordSet.BOF Then    
       'An empty recordset was returned
        QueryActiveDirectory = "Not Found"
    Else    'Records were found; loop through them
        While Not oRecordSet.EOF
            QueryActiveDirectory = oRecordSet.Fields("distinguishedName")
            oRecordSet.MoveNext
        Wend
    End If
End Function

Function CheckOutlookInstlled()
	Dim Wshshell
	Set WshShell = CreateObject("WScript.Shell")
	If WshShell.RegRead("HKCR\Outlook.Application\CurVer\")& "" = "" Then
	   wscript.echo "Outlook is not installed"
	   wscript.quit
	End If
End Function

Function CheckHasMailbox(UserID,DCServer)
	Dim objEnv
	Dim sUserLDAPName
	Dim objUser
	Dim strText
	Dim strTitle
	Dim Wrr
	Dim WshShell	
	
    Set WshShell = CreateObject("WScript.Shell")
	sUserLDAPName = QueryActiveDirectory(UserID)
	
	Set objUser = GetObject("LDAP://" & DCServer + "/" & sUserLDAPName)
	
	'check if user's mailbox has been created.
	If objUser.HomeMDB & ""  = "" Then
		strText="You have not mailbox yet now. Please Wait for mailbox to be created."
		strTitle="No Mailbox"
		Wrr=WshShell.Popup(strText,0,strTitle,vbokonly)
		WScript.Quit 1
	End If
End Function

Function GetExchangeServerName(UserID,DCServer)
    Dim objEnv
	Dim sUserLDAPName
	Dim objUser
	Dim ExchServer
	Dim ExchName
	Dim P1
	Dim C1
	Dim WshShell
	Set WshShell  = CreateObject("WScript.Shell")
		 
	sUserLDAPName = QueryActiveDirectory(UserID)
	Set objUser = GetObject("LDAP://" & DCServer + "/" & sUserLDAPName)
	ExchServer=objUser.HomeMDB
	P1=InStr(1,ExchServer,"CN=InformationStore")
	C1=InStr(1,ExchServer,"CN=Servers")-24-P1
	ExchName=Mid(ExchServer,P1+23,C1)
	GetExchangeServerName=ExchName
	
End Function

Function CheckHasNotesClient()
    Dim WshShell
    Dim NotesINIPath
    Set WshShell = CreateObject("WScript.Shell")
	NotesINIPath=WshShell.RegRead("HKLM\Software\Lotus\Notes\Path")     
	If NotesINIPath ="" Then
	   WScript.Echo("You should configure your notes client.")
	   WScript.Quit 1  
	End If  		
End Function 

Function ContactMigration(NamesFile,ExchServerName,UserID)
	Dim WshShell
	Dim objFSO
	Dim objConfFile
	Dim strCmdContact
	Dim str10
		
	Set WshShell = CreateObject("WScript.Shell")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
		
	'Change default mail program to Migrosoft Outlook.
	'WshShell.RegWrite "HKLM\SOFTWARE\Clients\Mail\","Microsoft Outlook","REG_SZ"
	
	'hangeClient()
	
	If objFSO.FileExists("c:\Migrate\tm\tmlog.txt") Then
	   objFSO.DeleteFile("c:\Migrate\tm\tmlog.txt")
	End If
		
		If objFSO.FileExists("c:\Migrate\tm\Contact.dat") Then
		   objFSO.DeleteFile("c:\Migrate\tm\Contact.dat")
		End If
	
	
	'Create configure file to convert Contact of Notes to Exchange 
	Set objConfFile = objFSO.OpenTextFile("c:\Migrate\TM\Contact.dat",8,True)
	objConfFile.Write "ConvertType=Addresses" & VbCrLf
	objConfFile.Write "ConvertFrom=NOTES" & VbCrLf
	objConfFile.Write "ConvertTo=Contacts" & VbCrLf
	objConfFile.Write "FromDatabase=" & NamesFile & VbCrLf
	objConfFile.Write "FromChecked=True" & VbCrLf
	objConfFile.Write "ToDatabase=" & ExchServerName & "!!" & UserID & VbCrLf
	objConfFile.Write "Convert=*" & VbCrLf
	objConfFile.Write "start" & VbCrLf
	objConfFile.Close
	
    strCmdContact="cmd /c c:\Migrate\TM\TMB8.EXE  /F C:\Migrate\TM\Contact.DAT /SL /Progress /L /S C:\Migrate\TM /d c:\Migrate /IM /NCFG "
	WshShell.run strCmdContact,8,True
	 	
End Function

Function GetContact()
	Dim regNotesDataPath
	Dim NamesFile
	Dim objFSO
	Dim WshShell
	Dim NotesINIPath
	Dim NotesINIFile
	Dim objFile
	Dim strLine
	Dim str1 
	Dim FindKeyFilename
	Dim FindMailFile
	Dim FindMailServer
	Dim NamesFilePath
	Dim strFind
	Dim strUserProfilePath
	Dim strMultiUser
	Dim objEnv
	
	Set WshShell = CreateObject("WScript.Shell")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	
	NotesINIPath = WshShell.RegRead("HKLM\Software\Lotus\Notes\Path")  
	'strMultiUser = WshShell.RegRead("HKLM\Software\Lotus\Notes\MultiUser") 
	Set objEnv = WshShell.Environment("PROCESS")
    strUserProfilePath = objEnv.item("USERPROFILE")
	
	If right(NotesINIPath,1) <> "\" Then
	   NotesINIPath=NotesINIPath & "\" 
	End If
	
'	If strMultiUser = 1 Then
'	  NotesINIFile = strUserProfilePath & "\" & "Local Settings\Application Data\Lotus\Notes\Data\" & "notes.ini"
'	Else   
	  NotesINIFile= NotesINIPath &  "notes.ini"  
'	End If     
	
	If Not objFSO.FileExists (NotesINIFile) Then
	   WScript.Echo("Can't find the file: " & NotesINIFile )
	   WScript.Quit 1
	End If
	
	Set objFile=objFSO.OpenTextFile(NotesINIFile,1)   'Open Notes.ini file for reading
	Do Until objFile.AtEndOfStream
	   strLine=objFile.ReadLine                       'Read notes.ini every line to check
	   str1=TRIM(strLine)
	   FindKeyFilename=InStr(1,str1,"Directory=") 
	   If FindKeyFilename = 1 Then
	      NamesFilePath=Right(str1,(Len(str1)-10))
	   End If
	Loop
	objFile.Close
	NamesFile= NamesFilePath & "\" & "Names.nsf"
	
	If Not objFSO.FileExists(NamesFile) Then
		WScript.Echo("cannot find the file:" & NamesFile)	
		WScript.Quit 1
	End If
	
	GetContact = NamesFile
	
End Function

Function CheckTransendMigrator()
	Dim objFSO
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FileExists("c:\Migrate\TM\TMB8.EXE") Then
	   WScript.Echo("Please copy Transend Migrator first.")
	   WScript.Quit 1
	End If
End Function

Function ChangeClient()
Dim WshShell
Dim Strcmd
Set WshShell = CreateObject("WScript.Shell")

Strcmd ="cmd /c c:\Migrate\ChangeToOutlookClient.exe"
WshShell.run Strcmd,8,True

End Function



