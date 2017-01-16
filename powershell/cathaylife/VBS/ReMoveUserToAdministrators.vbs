
strComputer = "."

Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators")

For Each objUser In objGroup.Members
    If objUser.Name <> "Administrator" AND objUser.Name <> "Domain Admins" AND objUser.Name <> "EADMS_AGENT" Then
        objGroup.Remove(objUser.AdsPath)
    End If
Next