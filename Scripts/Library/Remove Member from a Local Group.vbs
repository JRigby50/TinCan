Option Explicit 
Dim objGroup, strComputer, strMember
Dim objFSO, objServerList
Const ForReading = 1
strMember = "000-CarbonCopy-Admin"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objServerList = objFSO.OpenTextFile ("Servers.txt", ForReading)

Do until ObjServerList.AtEndOfStream
	strComputer = ObjServerList.ReadLine
	Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators,group")
	Call RmvUsrFrmGrp(objGroup,strMember)
Loop
 
Sub RmvUsrFrmGrp(objGroup,strMember) 
Dim objMember
On Error Resume Next
	For Each objMember In objGroup.Members
		If objMember.Name = strMember Then
		objGroup.Remove(objMember.ADsPath)
		objGroup.Setinfo
		End If 
	Next 
End Sub