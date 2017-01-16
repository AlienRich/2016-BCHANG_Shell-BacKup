'SearchAD.vbs
On Error Resume Next
' Connect to the LDAP server's root object
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")
strTarget = "LDAP://CN=Users,DC=cxldom09,DC=com,DC=tw"
'wscript.Echo "Starting search from " & strTarget

' Connect to Ad Provider
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCmd =   CreateObject("ADODB.Command")
Set objCmd.ActiveConnection = objConnection 

' Show only computers
'objCmd.CommandText = "SELECT Name, ADsPath FROM '" & strTarget & "' WHERE objectCategory = 'computer'"

' or show only users
objCmd.CommandText = "SELECT Name, ADsPath FROM '" & strTarget & "' WHERE objectCategory = 'user'"

' or show only groups
'objCmd.CommandText = "SELECT Name, ADsPath FROM '" & strTarget & "' WHERE objectCategory = 'group'"

Const ADS_SCOPE_SUBTREE = 2
objCmd.Properties("Page Size") = 100
objCmd.Properties("Timeout") = 30
objCmd.Properties("Searchscope") = ADS_SCOPE_SUBTREE
objCmd.Properties("Cache Results") = False

Set objRecordSet = objCmd.Execute

' Iterate through the results
objRecordSet.MoveFirst
Do Until objRecordSet.EOF
  
'====================
        If lcase(Left(objRecordSet.Fields("Name").Value,4)) = "0000" then
                Set objNewOU = GetObject("LDAP://OU=KS,OU=TM,OU=CathayBuilding,DC=cxldom09,DC=com,DC=tw")  
                objMoveUser = objNewOU.MoveHere ("LDAP://CN=" & objRecordSet.Fields("Name").Value &",DN=Users,DC=cxldom09,DC=com,DC=tw","CN=" & objRecordSet.Fields("Name").Value &"")  
        
        ElseIf lcase(Left(objRecordSet.Fields("Name").Value,4)) = "i464" then'
		Set objNewOU = GetObject("LDAP://OU=TC,OU=TM,OU=CathayBuilding,DC=cxldom09,DC=com,DC=tw")  
		objMoveUser = objNewOU.MoveHere ("LDAP://CN=" & objRecordSet.Fields("Name").Value &",CN=Users,DC=cxldom09,DC=com,DC=tw","CN=" & objRecordSet.Fields("Name").Value &"")  
                
        ElseIf lcase(Left(objRecordSet.Fields("Name").Value,4)) = "i461" then
		Set objNewOU = GetObject("LDAP://OU=TP,OU=TM,OU=CathayBuilding,DC=cxldom09,DC=com,DC=tw")  
		objMoveUser = objNewOU.MoveHere ("LDAP://CN=" & objRecordSet.Fields("Name").Value &",CN=Users,DC=cxldom09,DC=com,DC=tw","CN=" & objRecordSet.Fields("Name").Value &"") 

	ElseIf lcase(Left(objRecordSet.Fields("Name").Value,4)) = "i467" then'
		Set objNewOU = GetObject("LDAP://OU=KS,OU=TM,OU=CathayBuilding,DC=cxldom09,DC=com,DC=tw")  
		objMoveUser = objNewOU.MoveHere ("LDAP://CN=" & objRecordSet.Fields("Name").Value &",CN=Users,DC=cxldom09,DC=com,DC=tw","CN=" & objRecordSet.Fields("Name").Value &"")  



        
        End If

objRecordSet.MoveNext
Loop