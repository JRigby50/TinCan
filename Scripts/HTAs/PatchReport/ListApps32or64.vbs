On Error Resume Next
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
'Set objNetwork = CreateObject("Wscript.Network")
strServer = "LV48PCE00081501" 'objNetwork.ComputerName
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile(strServer & ".csv", True)

strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry1a = "DisplayName"
strEntry1b = "InstallDate"

Set objReg = GetObject("winmgmts://" & strServer & "/root/default:StdRegProv")
objReg.EnumKey HKLM, strKey, arrSubkeys
For Each strSubkey In arrSubkeys
	intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1)
	If intRet1 <> 0 Then
	'Next
	Else
    objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1b
		If strValue1 <> "" Then
			If InStr(strValue1, "KB") Or InStr(strValue1, "kb") Then
			objTextFile.WriteLine strServer & "," & strValue1 & "," & strValue1b
			End If
		End If
	End If
Next

strKey2 = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry2a = "DisplayName"
strEntry2b = "InstallDate"

Set objReg = GetObject("winmgmts://" & strServer & "/root/default:StdRegProv")
objReg.EnumKey HKLM, strKey2, arrSubkeys
For Each strSubkey In arrSubkeys
  intRet2 = objReg.GetStringValue(HKLM, strKey2 & strSubkey, strEntry2a, strValue2)
  If intRet2 <> 0 Then
  'Next
  Else
    objReg.GetStringValue HKLM, strKey & strSubkey, strEntry2b, strValue2b
		If strValue2 <> "" Then
			If InStr(strValue2, "KB") Or InStr(strValue2, "kb") Then
			objTextFile.WriteLine strServer & "," & strValue2 & "," & strValue2b
			End If
		End If
  End If
Next
objTextFile.Close

