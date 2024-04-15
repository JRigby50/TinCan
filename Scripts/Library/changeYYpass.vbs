Dim oldpass
Dim newpass
Dim yyuser
Const ADS_SCOPE_SUBTREE = 2

yyuser = InputBox("Enter yy account",, "yyaccount")
oldpass = InputBox("Enter Old Password",, "myoldpass")
newpass = InputBox("Enter New Password",, "mynewpass")

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

sqlstr = "SELECT distinguishedName FROM 'LDAP://dc=northgrum,dc=com' WHERE objectCategory='user' AND sAMAccountName='" & yyuser & "'"
objCommand.CommandText = sqlstr

Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst
Do Until objRecordSet.EOF
	DNstr = objRecordSet.Fields("distinguishedName").Value
    objRecordSet.MoveNext
Loop

Wscript.echo DNstr
	Set user = GetObject ("LDAP://" & DNstr)
	user.ChangePassword OldPass,newPass
wscript.echo "Exit with Err number: " & err.number