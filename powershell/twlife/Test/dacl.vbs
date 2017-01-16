Option Explicit

' Set Filesystem and Registry ACL
' ===============================
'
' Author: Tobias Oetiker <tobi@oetiker.ch> 
'         based on code by Nick Pearce, Craig Paterson and Rich Ellis
'
' Version: 1.2 -- 2001/05/08
'
' Changes: 1.1 Handle non-existing registry keys gracefully.
' Changes: 1.2 Added "K" option for simulating 1777 permission
'
' The purpose of this script is to allow ACL maniputlations
' to be performed by Microsoft Installer Packages (.msi)
'
' Add your ACL modification instructions to this script
' and integrate it as a Custom Action into the MSI.
'
' SETUP
' =====
'
' Note aclfix needs ADsSecurity.dll and RegObj.dll to work. 
' You can get ADsSecurity from ADSI SDK 2.5 under (/ResourceKit/ADsSecurity.dll)
'
' Get the sdk from
' http://msdownload.microsoft.com:80/msdownload/adsi/2.5/sdk/x86/en/Sdk.zip
' copy the dll to a place in your path and run 
'
' RegObj.dll is included in Office 2000 SR2 or also available directly from
' MS for registered VB users.
'
' regsvr32 regobj.dll
' regsvr32 adssecurity.dll 
'
' Usage with WISE for Windows Installer
' =====================================
' * add a copy of the two dlls to the package and install them somewhere
'   below the INSTALLDIR of the package. Make sure you click Self register in
'   the file property dialog
'
' * customize the dacl.vbs according to the needs of the application and add it
'   to the msi somewhere below INSTALLDIR. Maybe next to the dlls
' 
' * Add a custom action: Type:    Call Exe File
'                        Source:  File on destination machine
'                        Name:    DACL
'                        InstDir: SystemFolder
'                        Exe Cmd: wscript.exe "[!dacl.vbs]"  
'                        Sequenc: Install Execute Sequence (Before InstallFinalize)'                        
'			 Condition: NOT REMOVE~="ALL"
'                        I-S Opt: System Context
'                        Process: Asynch, Wait at end of sequence
'
' TODO
' ====
' Using Add with the Registry creates working results, but somehow the ACEs are
' not in the proper order, and regedt32 complains when you look at them ...
' it also seems that maybe I am not adding all the reg entries necessary ...
' lack of docu ... sorry ... help appreciated. I guess it has something todo
' with inheritance ... 
'
' USAGE
' -----
'
' DACL function, url, "ace, ace, ..."
'
' function -- Add, Rm, Set
'
' url -- FILE://....       change this File/Folder
'        FILE://c:\home\   change this Folder and everything below
'        FILE://c:\home\\  change this Folder and Folders below
'        RGY://\HKEY_LOCAL_MACHINE\SOFTWARE    change this property
'        RGY://...\  and RGY://\...\\ are the same as indiviual
'                    registry values have no acls assigned
'
' ace -- account:rights
'
' account -- user or group
'
' rights (file) --  F - Full, C - Change, R - Read + Execute, 
'                   S - Read + Write + Execute, L - List
'                   K - Special Case, like 1777. Create file and work with your own only
' rights (registry) --  F - Full, R - Read
'
' EXAMPLES
' --------
' DACL "Add", "FILE://w:\hello.txt", "users:F,moetiker:F"
' DACL "Add", "FILE://w:\hello\",    "users:R,oetiker:F,moetiker:F"
' DACL "Rm",  "FILE://w:\oops.txt",  "user"   'remove whatever is there under user
' DACL "Add",  "FILE://w:\oops.txt",  "user:F" ' add Full control for user
' DACL "Add", "RGY://\HKEY_CURRENT_USER\SOFTWARE\ipswitch\ws_ftp\", "users:F"
'=============================================================================
DumpAcl "FILE://C:"
DACL "Rm", "FILE://C:", "users,CREATOR OWNER"
DACL "Add", "FILE://C:", "users:R"
DACL "Add", "FILE://C:", "users:K"
DumpAcl "FILE://C:"
'=============================================================================
' Implementation
' --------------
Sub DumpAcl (url)
    Dim sd, dacl, ace, sec
    Print url
    Set sec = Wscript.CreateObject("ADsSecurity")
    Set sd = sec.GetSecurityDescriptor( CStr(url) )
    Set dacl = sd.DiscretionaryAcl
    For Each ace In dacl
        Print "   " & ace.trustee & _
              " Type: " & ace.AceType & _
              " Mask: " & ace.AccessMask & _
              " AceFlags: " & ace.AceFlags & _
              " Flags: " & ace.Flags & _
	      " OType: " & ace.ObjectType & _
	      " IOTyp: " & ace.InheritedObjectType

    Next

    set sec = Nothing
    set sd = Nothing
    set dacl = Nothing
    
End Sub

Sub AclEdit( action, url, acl, UType )
    Dim acls, dacl, dummy, sec, sd, ace, acea, usera, user, perm, acsplit
    Const ADS_ACETYPE_ACCESS_ALLOWED = 0
    Const ADS_ACETYPE_ACCESS_DENIED = 1
    Const ADS_ACEFLAG_INHERIT_ACE = 2
    Const ADS_ACEFLAG_INHERIT_ONLY_ACE         = 8
    Const ADS_ACEFLAG_SUB_NEW = 9

    Print "Edit: " & action & " " & url & " " & acl & " " & utype
    acls = split(acl,",")
    
    Set sec = Wscript.CreateObject("ADsSecurity")
    on error resume Next

    ' without cstr this will break ... !!!
    Set sd = sec.GetSecurityDescriptor( CStr(url) )

    If ErrHandler("Get SD for " & url ) Then
        On Error GoTo 0
        Exit Sub
    End If

    Set dacl = sd.DiscretionaryAcl
    dummy = dacl.AceCount ' this will throw an error if there is no DACL    
    If ErrHandler("Get DACL for " & url ) Then
        On Error GoTo 0
        Exit Sub
    End If

    ' DumpAcl url

'    If action = "Rm" Or action = "Add" Then
    If action = "Rm" Then
        ' for Add we remove the ACEs for the folks for whoom need new ones
        For Each ace In dacl
            acea = split (LCase(ace.trustee & "\" & ace.trustee),"\")
            If acea(0) <> "nt authority" Then
                For Each user In acls
                    usera = split (LCase(user),":")        
                    If acea(1) = usera(0) Then
                        Print "Remove ACE: " & ace.trustee
                        dacl.RemoveACE ace                    
                        ErrHandler("Remove ACE for " & ace.trustee & _
                                   " from " & url)
                    End If
                Next
            End if
        Next

    ElseIf action = "Set" Then
        For Each ace In dacl
            acea = split (LCase(ace.trustee & "\" & ace.trustee),"\")
            If acea(0) <> "nt authority" Then
                dacl.RemoveACE ace
                ErrHandler("Remove ACE for " & ace.Trustee & " from " & url)
                Print "Remove ACE: " & ace.trustee
            End if
        Next    
    Elseif action <> "Add" Then
        Wscript.Echo "Unknown Action: " & action
    End If
    
    If action = "Set" Or action = "Add" Then
        For Each dummy In acls
            acsplit = split (dummy,":")
            user = acsplit(0)
            perm = acsplit(1)
            Print action & " " & utype & " " & user & " " & perm
            Select Case UType
                Case "DIRECTORY"
                    ' folders require 2 aces for user (to do with inheritance)
                    AddFileAce dacl, user, perm, _
                               ADS_ACETYPE_ACCESS_ALLOWED, _
                               ADS_ACEFLAG_SUB_NEW
		    if ucase(perm) <> "K" then
                        AddFileAce dacl, user, perm, _
                               ADS_ACETYPE_ACCESS_ALLOWED, _
                               ADS_ACEFLAG_INHERIT_ACE
                    end if
                case "FILE"
                    AddFileAce dacl, user, perm, _
                               ADS_ACETYPE_ACCESS_ALLOWED,0
                case "REGISTRY"
                    AddRegAce dacl, user, perm, _
                              ADS_ACETYPE_ACCESS_ALLOWED, _
                              ADS_ACEFLAG_INHERIT_ACE
            End Select
        Next

    End If

    
    sd.DiscretionaryAcl = dacl
    If ErrHandler("Get SD for " & url ) Then    
        On Error GoTo 0
        Exit Sub
    End If

    sec.SetSecurityDescriptor sd    
   
    If ErrHandler("Get SD for " & url ) Then
        On Error GoTo 0
        Exit Sub
    End If
    
    Set sd = Nothing
    Set dacl = Nothing
    Set sec = Nothing

    ' DumpAcl url

    On Error GoTo 0  
End Sub

Sub AddRegACE(dacl, user, perm , acetype, aceflags)
    ' Add registry ACE
    Dim ace
    
    Const ADS_ACETYPE_ACCESS_ALLOWED = 0
    Const RIGHT_REG_READ = &H20019
    Const RIGHT_REG_FULL = &HF003F

    
    Set ace = CreateObject("AccessControlEntry")
    ace.Trustee = user
    
    Select Case UCase(perm)
        ' specified rights so far only include FC & R. Could be expanded though
        Case "F"
            ace.AccessMask = RIGHT_REG_FULL
        Case "R"
            ace.AccessMask = RIGHT_REG_READ
    End Select
    
    ace.AceType = acetype
    ace.AceFlags = aceflags
    dacl.AddAce ace
    ErrHandler("Add Ace for " & user )

    set ace=Nothing

End Sub

Sub AddFileAce(dacl,user, perm, acetype, aceflags)
    ' add ace to the specified dacl
    Dim ace
    
    Const RIGHT_LIST = &H4
    Const RIGHT_READ = &H80000000
    Const RIGHT_EXECUTE = &H20000000
    Const RIGHT_WRITE = &H40000000
    Const RIGHT_DELETE = &H10000
    Const RIGHT_FULL = 2032127 ' &H10000000
    Const RIGHT_CHANGE_PERMS = &H40000
    Const RIGHT_TAKE_OWNERSHIP = &H80000

    
    Set ace = CreateObject("AccessControlEntry")
    ace.Trustee = user
    
    select case ucase(perm)
        ' specified rights so far only include FC & R. Could be expanded though
        case "F"
            ace.AccessMask = RIGHT_FULL
        case "C"
            ace.AccessMask = RIGHT_READ or RIGHT_WRITE Or _
               RIGHT_EXECUTE or RIGHT_DELETE
        case "R"
            ace.AccessMask = RIGHT_READ or RIGHT_EXECUTE
        case "S" 'Special
            ace.AccessMask = RIGHT_READ or RIGHT_WRITE or RIGHT_EXECUTE
        case "L" 'List
            ace.AccessMask = RIGHT_LIST
        case "K" ' User can Create and work files in the dir specified and nothing else
            ace.AceType = 0
            ace.AccessMask = 1048582 ' create files
	    ace.AceFlags = 0
            dacl.AddAce ace
            set ace=Nothing
            Set ace = CreateObject("AccessControlEntry")    
            ace.Trustee = "CREATOR OWNER"
            ace.AccessMask = RIGHT_FULL
            acetype=0
            aceflags=9
    end select
    
    ace.AceType = acetype
    ace.AceFlags = aceflags
    dacl.AddAce ace
    ErrHandler("Add Ace for " & user )

    set ace=Nothing

End Sub
   
Sub DACL(action,url,acl)
    Dim argarry, utype, upath, walk, ftype, fs, rfldr, file, sfldr
    Dim ro, rk, regval,skey
    argarry = split(url,"://")
    utype = argarry(0)
    upath = argarry(1)

    Print "Action: " & action & " " & utype & "--" & upath & " " & acl

    If Right(upath,2) = "\\" Then
        walk = "\\" ' folders only 
        upath = Left(upath, Len(upath)-2)
    ElseIf Right(upath,1) = "\" Then
        walk = "\" ' files and folders
        upath = left(upath, len(upath)-1)      
    End If
    
    If utype = "FILE" Then
        
        Set fs=Wscript.CreateObject("Scripting.FileSystemObject")
        Print "---" & upath
        If fs.FileExists(upath) Then
            Set rfldr=fs.GetFile(upath)
            ftype = "FILE" 'directory
        ElseIf fs.FolderExists(upath) Then
            Set rfldr=fs.GetFolder(upath)
            ftype = "DIRECTORY" 'file
        Else
            ' its neither file nor folder ... maybe it does not exist ...
            wscript.echo "Can't find " & upath
            Exit Sub
        End If
        
        AclEdit action, "FILE://" & rfldr.path, acl, ftype

        If ftype = "FILE" Then 'if this is a file our work is done
            Exit Sub
        End If
        
        If walk = "\" then
            For Each file In rfldr.files
                AclEdit action, "FILE://" & file , acl, "FILE"
            Next
        End If
        
        if walk = "\" or walk = "\\" then 
            for each sfldr in rfldr.subfolders
                DACL action, "FILE://" & sfldr & walk, acl
            next
        end if
        
    elseif utype = "RGY" Then

        Set ro = CreateObject("RegObj.Registry")
        on error resume Next
        Set rk = ro.RegKeyFromString( upath )
        If ErrHandler("Get Registry Key " & upath ) Then
            On Error GoTo 0
            Exit Sub
        End If
        
        AclEdit action, "RGY://" & rk.FullName, acl,"REGISTRY"  
                        
        if walk = "\" or walk = "\\" then
            for each skey in rk.Subkeys
                DACL action,"RGY://" &  skey.FullName & walk, acl
            next
        end If
        
    else
        Wscript.Echo "Unsupported URL Type: " & utype
    end If
    On Error GoTo 0
    
End Sub

Function ErrHandler(what)
    If Err.Number > 0 Then
        WScript.Echo what & " Error " & Err.Number & ": " & Err.Description
        Err.Clear
        Return True
    End If
    ErrHandler = False
End Function

Sub Print(Str)
    'strip when debugging
    'wscript.echo Str
End Sub






