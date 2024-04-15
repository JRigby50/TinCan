On Error Resume Next 

Set objGroup = GetObject _
  ("LDAP://cn=Patch-Auto-East-Any-0200,ou=ServerPatching,ou=Groups,dc=northgrum,dc=com")
objGroup.GetInfo

arrMemberOf = objGroup.GetEx("member")

WScript.Echo "Members:"
For Each strMember in arrMemberOf
    WScript.echo strMember
Next


Dim strDomain, strBase, strScope, strAttributes, strSort, strADFaxNumber, arrTemp
strDomain = "DC=northgrum,DC=com"
strBase =  "<LDAP://" & strDomain & ">;"
strScope = "SubTree "
strSort = "CN" 'CanonicalName
strAttributes = "objectClass,objectCategory,distinguishedName,cn,displayName,mail,telephoneNumber;"
