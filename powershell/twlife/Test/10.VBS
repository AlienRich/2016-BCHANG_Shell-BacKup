Option Explicit

Dim objRootDSE, strDNSDomain

Set objRootDSE = GetObject(LDAP://RootDSE)
strDNSDomain = objRootDSE.Get("defaultNamingContext")
Select Case LCase(strDNSDomain)
Case "dc=domain1,dc=com"
	' .... code to map drive and printer.
Case "dc=domain2,dc=com"
	' .... code to map drive and printer.
End Select