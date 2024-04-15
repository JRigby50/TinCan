ADS_CHASE_REFERRALS_SUBORDINATE = &h20 
StrMail = InputBox("Enter Email address")' 
dtStart = TimeValue(Now()) 
Set objConnection = CreateObject("ADODB.Connection") 
objConnection.Open "Provider=ADsDSOObject;" 
  
Set objCommand = CreateObject("ADODB.Command") 
objCommand.ActiveConnection = objConnection 
  
objCommand.Properties("Chase Referrals") = _  
   ADS_CHASE_REFERRALS_SUBORDINATE 
  
objCommand.CommandText = _ 
    "<LDAP://dc=northgrum,dc=com>;(&(objectCategory=person)(objectClass=user)" & _ 
    "(proxyAddresses=" & "SMTP:" & StrMail & "));employeeid,proxyaddresses,name;subtree" 
        
Set objRecordSet = objCommand.Execute 
  
If objRecordset.RecordCount = 0 Then 
    WScript.Echo "The Email address: " & StrMail & " does not exist." 
Else 
    WScript.Echo "The Email address was found and belongs to user "  & objRecordset.fields("name") 
     
End If 
  
objConnection.Close 
