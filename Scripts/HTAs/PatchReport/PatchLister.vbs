Option Explicit
Dim strServer, strThisMonth, strThisYear, objFSO, objTextFile, bolPatch
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
strServer = "LV48PCE00081501"
strThisMonth = "04"
strThisYear = "2011"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile(strServer & ".csv", True)

ListPatches strServer

Sub ListPatches(strServer)
'On Error Resume Next
Dim intRet1, intRet2
Dim arrTmp, arrSubkeys
Dim objReg
Dim strKey, strEntry1a, strEntry1b, strKey2, strEntry2a, strEntry2b, strKB, strDay, strMonth, strYear, strSubkey
Dim strValue1, strValue2, strValue1b, strValue2b
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry1a = "DisplayName"
strEntry1b = "InstallDate"
strKey2 = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry2a = "DisplayName"
strEntry2b = "InstallDate"

Set objReg = GetObject("winmgmts://" & strServer & "/root/default:StdRegProv")
objReg.EnumKey HKLM, strKey, arrSubkeys
For Each strSubkey In arrSubkeys
	intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1)
	If intRet1 <> 0 Then
	'Next
	Else
    objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1b
		If strValue1 <> "" And strValue1b <> "" Then
			If InStr(strValue1, "(KB") Or InStr(strValue1, "(kb") Then  '- KB
			arrTmp = Split(strValue1,"(")
			strKB=arrTmp(1)
			strKB=Left(strKB,Len(strKB)-1)
			strDay = Right(strValue1b,2)
			strMonth = Mid(strValue1b,5,2)
			strYear = Left(strValue1b,4)
				If strMonth = strThisMonth AND strYear = strThisYear Then
				bolPatch = True
				objTextFile.WriteLine strServer & "," & strKB & "," & strMonth & "/" & strDay & "/" & strYear
				End If
			End If
			If InStr(strValue1, "- KB") Or InStr(strValue1, "- kb") Then
			arrTmp = Split(strValue1,"-")
			strKB=arrTmp(1)
			strKB=Trim(strKB)
			strDay = Right(strValue1b,2)
			strMonth = Mid(strValue1b,5,2)
			strYear = Left(strValue1b,4)
				If strMonth = strThisMonth AND strYear = strThisYear Then
				bolPatch = True
				objTextFile.WriteLine strServer & "," & strKB & "," & strMonth & "/" & strDay & "/" & strYear
				End If
			End If
		End If
	End If
Next
objReg.EnumKey HKLM, strKey2, arrSubkeys
If arrSubkeys Then 
	For Each strSubkey In arrSubkeys
	  intRet2 = objReg.GetStringValue(HKLM, strKey2 & strSubkey, strEntry2a, strValue2)
	  If intRet2 <> 0 Then
	  'Next
	  Else
	    objReg.GetStringValue HKLM, strKey & strSubkey, strEntry2b, strValue2b
			If strValue2 <> ""  And strValue2b <> "" Then
				If InStr(strValue2, "(KB") Or InStr(strValue2, "(kb") Then
				arrTmp = Split(strValue2,"(")
				strKB=arrTmp(1)
				strKB=Left(strKB,Len(strKB)-1)
				strDay = Right(strValue2b,2)
				strMonth = Mid(strValue2b,5,2)
				strYear = Left(strValue2b,4)
					If strMonth = strThisMonth AND strYear = strThisYear Then
					bolPatch = True
					objTextFile.WriteLine strServer & "," & strKB & "," & strMonth & "/" & strDay & "/" & strYear
					End If
				End If
				If InStr(strValue2, "- KB") Or InStr(strValue2, "- kb") Then
				arrTmp = Split(strValue2,"-")
				strKB=arrTmp(1)
				strKB=Trim(strKB)
				strDay = Right(strValue2b,2)
				strMonth = Mid(strValue2b,5,2)
				strYear = Left(strValue2b,4)
					If strMonth = strThisMonth AND strYear = strThisYear Then
					bolPatch = True
					objTextFile.WriteLine strServer & "," & strKB & "," & strMonth & "/" & strDay & "/" & strYear
					End If
				End If
			End If
	  End If
	Next
End If
End Sub
objTextFile.Close

