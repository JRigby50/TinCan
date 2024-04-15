Option Explicit 
Dim objGroup, strComputer 
 
strComputer = "lv48pce00081501"
 
Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators,group") 
Wscript.Echo "Members of local Administrators group on computer " & strComputer 

Call EnumGroup(objGroup) 
 
Sub EnumGroup(objGroup) 
Dim objMember
On Error Resume Next
For Each objMember In objGroup.Members 
	If objMember.Name = "000-CarbonCopy-Admin" Then
	Wscript.Echo strOffset & objMember.Name & " (" & objMember.Class & ")" 
	End If 
Next 
End Sub