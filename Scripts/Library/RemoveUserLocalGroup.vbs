Option Explicit
Dim strComputer, objGroup, objUser
strComputer = "lv48pce00081501"
Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators,group")
'For Each objUser In  objGroup.Members
'WScript.Echo objUser.name
Set objUser = GetObject("WinNT://" & strComputer & "/000-CarbonCopy-Admin,group")

'objGroup.Remove(objUser.ADsPath)
'Next
MsgBox "End"
