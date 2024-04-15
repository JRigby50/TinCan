'*****One*****
'Set objEmail = CreateObject("CDO.Message")
'objEmail.From = "netsysnotif@ngc.com"
'objEmail.To = "james.rigby@ngc.com"
'objEmail.Subject = "Test1" 
'objEmail.Textbody = "Test1 Test1 Test1."
'objEmail.Send
'*****Two*****
Set objEmail = CreateObject("CDO.Message")
objEmail.From = "netsysnotif@ngc.com"
objEmail.To = "james.rigby@ngc.com"
objEmail.Subject = "Test2" 
objEmail.Textbody = "Test2 Test2 Test2."
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "relaymd001.northgrum.com" 
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update
objEmail.Send
