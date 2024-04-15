'Define variables, constants and objects

strComputer="."
Const HKEY_USERS = &H80000003
Set objWbem = GetObject("winmgmts:")
Set objRegistry = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")
Set objWMIService = GetObject("winmgmts:"  & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

'Go and get the currently logged on user by checking the owner of the Explorer.exe process.  

Set colProc = objWmiService.ExecQuery("Select Name from Win32_Process" & " Where Name='explorer.exe' and SessionID=0")

If colProc.Count > 0 Then
For Each oProcess In colProc
oProcess.GetOwner sUser, sDomain
Next
End If

'Loop through the HKEY_USERS hive until (ignoring the .DEFAULT and _CLASSES trees) until we find the tree that 
'corresponds to the currently logged on user.
lngRtn = objRegistry.EnumKey(HKEY_USERS, "", arrRegKeys)

For Each strKey In arrRegKeys
If UCase(strKey) = ".DEFAULT" Or UCase(Right(strKey, 8)) = "_CLASSES" Then
Else

Set objSID = objWbem.Get("Win32_SID.SID='" & strKey & "'")

'If the account name of the current sid we're checking matches the accountname we're looking for Then
'enumerate the Network subtree
If objSID.accountname = sUser Then 
regpath2enumerate = strkey & "\Network" 'strkey is the SID
objRegistry.enumkey hkey_users, regpath2enumerate, arrkeynames

'If the array has elements, go and get the drives info from the registry
If Not (IsEmpty(arrkeynames)) Then
For Each subkey In arrkeynames
regpath = strkey & "\Network\" & subkey
regentry = "RemotePath"
objRegistry.getstringvalue hkey_users, regpath, regentry, dapath
wscript.echo subkey & ":" & vbTab & dapath
Next
End If
End If
End If
Next

