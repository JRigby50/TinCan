Const adVarChar = 200
Const MaxCharacters = 255
Const ForReading = 1
Const ForWriting = 2
Const ADS_SCOPE_SUBTREE = 2
Const FOR_WRITING = 2

'************************ Create Server List ****************************

set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists("Serverlist.txt") Then
	objFSO.DeleteFile("Serverlist.txt")
End IF
Set objFile = objFSO.CreateTextFile("Serverlist.txt", FOR_WRITING)
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCOmmand.ActiveConnection = objConnection
Set RootDse = GetObject( "LDAP://RootDse" )
strADSPath = "LDAP://" & RootDse.get( "DefaultNamingContext" )
objCommand.CommandText = _
	    "Select Name, Location from '" & strADSPath &_
		"' Where objectClass='computer'"
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    objFile.WriteLine objRecordSet.Fields("Name").Value
    objRecordSet.MoveNext
Loop



'***********************  Sort Server List  ******************************

Set DataList = CreateObject("ADOR.Recordset")
DataList.Fields.Append "ComputerName", adVarChar, MaxCharacters
DataList.Open

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("serverlist.txt", ForReading)

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    DataList.AddNew
    DataList("ComputerName") = strLine
    DataList.Update
Loop

objFile.Close

DataList.Sort = "ComputerName"

DataList.MoveFirst
Do Until DataList.EOF
    strText = strText & DataList.Fields.Item("ComputerName") & vbCrLf
    DataList.MoveNext
Loop

Set objFile = objFSO.OpenTextFile("serverlist.txt", ForWriting)

objFile.WriteLine strText
objFile.Close

