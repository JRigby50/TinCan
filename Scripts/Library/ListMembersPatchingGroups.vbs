Option Explicit
'Define Constants And Variables
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Dim objFSO, objGroupList, objGroupMembersList, objGroup
Dim strGroupName, strNow, strMember, arrMemberOf
strNow = Now
'Create Basic Objects
Set objFSO = WScript.CreateObject("Scripting.Filesystemobject")
Set objGroupList = objFSO.OpenTextFile("PatchGroups.txt", ForReading)

If objFSO.FileExists("GroupMembers.txt") Then
objFSO.DeleteFile("GroupMembers.txt")
End If
Set objGroupMembersList = objFSO.CreateTextFile("GroupMembers.txt", ForWriting)
'*****
Do Until objGroupList.AtEndOfStream
strGroupName = objGroupList.ReadLine
objGroupMembersList.WriteLine strGroupName & " Members:"
objGroupMembersList.WriteBlankLines(1)
On Error Resume Next 
Set objGroup = GetObject("LDAP://cn=" & strGroupName & ",ou=ServerPatching,ou=Groups,dc=northgrum,dc=com")
objGroup.GetInfo
arrMemberOf = objGroup.GetEx("member")
'objGroupMembersList.WriteLine "Members:"
For Each strMember in arrMemberOf
    objGroupMembersList.WriteLine strMember
Next
objGroupMembersList.WriteBlankLines(2)
'*****
Loop
objGroupList.Close
objGroupMembersList.Close