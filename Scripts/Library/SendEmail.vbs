Option Explicit
Dim strEmailAddress, errEmail
strEmailAddress = "james.rigby@ngc.com"

errEmail = SendEmail(strEmailAddress)

Function SendEmail(strEmailAddress)
Dim strMessage, objEmail
strMessage = strTextBody
On Error Resume Next
Set objEmail = CreateObject("CDO.Message")
objEmail.From = "PrintServerBackup@ngc.com"
objEmail.To = strEmailAddress
objEmail.Subject = "Printer Backup Log"
objEmail.TextBody = strMessage
objEmail.AddAttachment ("C:\IT\MTLog.txt")
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "relayeast.northgrum.com" 
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update
objEmail.Send
SendEmail = Err
End Function