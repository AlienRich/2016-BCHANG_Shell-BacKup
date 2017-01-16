OPTION EXPLICIT

DIM strFilter, strRoot, strScope, strGroupName
DIM strNETBIOSDomain, strGroupDN
DIM cmd, rs,cn, objGroup
DIM objSystemInfo
' ********************* Setup *********************

' Default filter for all user accounts (ammend if required)
strFilter = "(&(objectCategory=person)(objectClass=user))"
' scope of search (default is subtree - search all child OUs)
strScope = "subtree"

' search root. e.g. ou=MyUsers,dc=wisesoft,dc=co,dc=uk
strRoot = "OU=tp,OU=tm,OU=CathayBuilding,dc=cxldom09,dc=com,dc=tw"

' Group to add
strGroupName = "TM"

' *************************************************

SET objSystemInfo = CREATEOBJECT("ADSystemInfo") 
'Required for name translate
strNETBIOSDomain = objSystemInfo.DomainShortName

' Convert group name to distinguished name
strGroupDN =  GetDN(strNETBIOSDomain,strGroupName)

SET objGroup = GETOBJECT("LDAP://" & strGroupDN)
SET cmd = CREATEOBJECT("ADODB.Command")
SET cn = CREATEOBJECT("ADODB.Connection")
SET rs = CREATEOBJECT("ADODB.Recordset")

cn.open "Provider=ADsDSOObject;"
cmd.activeconnection = cn

cmd.commandtext = "<LDAP://" & strRoot & ">;" & strFilter & ";ADsPath,sAMAccountName;" & strScope

'**** Bypass 1000 record limitation ****
cmd.properties("page size")=1000

SET rs = cmd.EXECUTE

' for each item returned by the Active Directory query
WHILE rs.eof <> TRUE AND rs.bof <> TRUE

	ON ERROR RESUME NEXT
	' Add user to group
	objGroup.Add rs("ADsPath")

	' Error handling (user might already be a member of the group)
	IF err.number = -2147019886 THEN
		wscript "User '" & rs("sAMAccountName") & _
			"' is already a member of the '" & objGroup.sAMAccountName & "' group"
	ELSEIF err.number <> 0 THEN
		wscript "Error: " & err.number & " adding user '" & _
			 rs("sAMAccountName") & "' to the '" & objGroup.sAMAccountName & "' group"
	ELSE
		wscript "Added user '" & rs("sAMAccountName") & "' to the " & _
			objGroup.sAMAccountName & "' group"
	END IF
	err.clear

	ON ERROR GOTO 0

	rs.movenext
WEND

' Close ADO connection
cn.close

' Function to convert name into distinguished name format
FUNCTION GetDN(BYVAL strDomain,strObject)
	' Use name translate to return the distinguished name
	' of a user from the NT UserName (sAMAccountName)
	' and the NETBIOS domain name.
	DIM objTrans

	SET objTrans = CREATEOBJECT("NameTranslate")
	objTrans.Init 1, strDomain
	objTrans.SET 3, strDomain & "\" & strObject
	GetDN = objTrans.GET(1) 

END FUNCTION