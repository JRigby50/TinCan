on error resume next

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
Set objNetwork = CreateObject("Wscript.Network")
strComputer = objNetwork.ComputerName

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile("c:\" & strcomputer & ".txt", True)

strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry1a = "DisplayName"
strEntry1b = "QuietDisplayName"

Set objReg = GetObject("winmgmts://" & strComputer & _
 "/root/default:StdRegProv")
objReg.EnumKey HKLM, strKey, arrSubkeys

For Each strSubkey In arrSubkeys
  intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, _
   strEntry1a, strValue1)
  If intRet1 <> 0 Then
    objReg.GetStringValue HKLM, strKey & strSubkey, _
     strEntry1b, strValue1
  End If
  If strValue1 <> "" Then
objTextFile.WriteLine strValue1 '& "¿ " & strcomputer & "¿"
  End If
Next

strKey2 = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry2a = "DisplayName"
strEntry2b = "QuietDisplayName"

Set objReg = GetObject("winmgmts://" & strComputer & _
 "/root/default:StdRegProv")
objReg.EnumKey HKLM, strKey2, arrSubkeys

For Each strSubkey In arrSubkeys
  intRet2 = objReg.GetStringValue(HKLM, strKey2 & strSubkey, _
   strEntry2a, strValue2)
  If intRet2 <> 0 Then
    objReg.GetStringValue HKLM, strKey & strSubkey, _
     strEntry2b, strValue2
  End If
  If strValue2 <> "" Then
objTextFile.WriteLine strValue2 '& "¿ " & strcomputer & "¿"
  End If
Next

objTextFile.Close

