

'Sub Installed_Font_Report
  'Creates Report of Fonts
'  Const FONTS = &H14&
'Set objShell = CreateObject("Shell.Application")
'Set objFolder = objShell.Namespace(FONTS)
'Set objFolderItem = objFolder.Self
'Wscript.Echo objFolderItem.Path

'Set colItems = objFolder.Items
'For Each objItem in colItems
'    Wscript.Echo objItem.Name
'Next
	
On error resume next
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
Set objNetwork = CreateObject("Wscript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strComputer = objNetwork.ComputerName
Set objTextFile = objFSO.CreateTextFile("c:\" & strcomputer & " Installed Fonts.txt", True)
objTextFile.WriteLine "Installed Fonts"
objTextFile.WriteLine
strKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
' HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts
strValueName = "Fonts"
strEntry1a = "DisplayName"
strEntry1b = "QuietDisplayName"
' objReg.GetDWORDValue
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
'objReg.Item HKLM, strKey, arrFonts
'objReg.GetDWORDValue HKLM, strKey, arrFonts
'objReg.RegRead HKLM, strKey, arrFonts
'objReg.EnumKey HKLM, strKey, arrFonts
objReg.EnumValues HKLM, strKey, arrFonts
'objReg.GetStringValue HKLM, strKey, strValueName, arrFonts
For Each strSubkey In arrFonts
  intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1)
  If intRet1 <> 0 Then
    objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1
  End If
  
  If strValue1 <> "" Then
objTextFile.WriteLine strValue1
  End If
Next

'strKey2 = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
'strEntry2a = "DisplayName"
'strEntry2b = "QuietDisplayName"

'Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")
'objReg.EnumKey HKLM, strKey2, arrFonts

'For Each strSubkey In arrFonts
'  intRet2 = objReg.GetStringValue(HKLM, strKey2 & strSubkey, strEntry2a, strValue2)
'  If intRet2 <> 0 Then
'    objReg.GetStringValue HKLM, strKey & strSubkey, strEntry2b, strValue2
'  End If
'  If strValue2 <> "" Then
'objTextFile.WriteLine strValue2
'  End If
'Next
objTextFile.Close
'End Sub