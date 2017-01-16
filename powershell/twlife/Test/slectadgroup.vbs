Set objADSysInfo = CreateObject(\"ADSystemInfo\") 
strUser = objADSysInfo.UserName 

S = strUSer & vbCrLf & \"屬於以下群組:\" & vbCrLf & \"==============================\" & vbCrLf 
Set objUser = GetObject(\"LDAP://\" & strUser) 

If IsArray(objUser.MemberOf) Then 
   For Each group in objUser.MemberOf 
       Set objGroup = GetObject(\"LDAP://\" & group) 
       S = S & objGroup.CN & vbCrLf 
   Next 
Else 
   S = S & objUser.MemberOf 
End If 

WScript.Echo S