On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2

strUser=InputBox("請輸入被鎖定的AD使用者帳號：","訊息視窗")

Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCommand =   CreateObject("ADODB.Command")
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
objCommand.CommandText = "SELECT ADsPath FROM 'LDAP://dc=test,dc=com' WHERE NAME=' " & strUser & "'"

Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst
Do Until objRecordSet.EOF
     ldap_path=objRecordSet.Fields("ADsPath").Value
     Set objUser=GetObject(ldap_path)
     objUser.IsAccountLocked=False
     objUser.SetInfo
     wscript.echo "OK!"
     objRecordSet.MoveNext
Loop

