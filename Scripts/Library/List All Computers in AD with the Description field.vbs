'List All Computers in AD with the Description field

Dim strDescription

Const ADS_SCOPE_SUBTREE = 1000

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCOmmand.ActiveConnection = objConnection
objCommand.CommandText = _
    "Select Name, Location, Description from 'LDAP://ehsd.east-haven.k12.ct.us/DC=ehsd,DC=east-haven,DC=k12,DC=ct,DC=us' Where objectClass='computer'"
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Wscript.Echo "Computer Name, Computer Description"

Do Until objRecordSet.EOF

                strDescription = objRecordSet.Fields("description").Value

                If IsNull (strDescription) Then
                                WScript.Echo objRecordSet.Fields("Name").Value
                Else
                                WScript.Echo objRecordSet.Fields("Name").Value & " ,  " & strDescription(0)
                End If

    objRecordSet.MoveNext
Loop

