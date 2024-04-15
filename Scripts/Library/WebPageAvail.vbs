'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME: Web Page Monitor
'
' AUTHOR: Josiah Woodhouse - http://woodhouseworld.spaces.live.com/
' DATE  : 8/10/2006
'
' COMMENT: Monitors web Pages
'	I like this script because it can be changed to handle various scenarios
'==========================================================================
' Add or change the pages you want to monitor
MyUrls = Array("http://www.adventureworks.com", "http://www.adventureworks.com/Page1.aspx", "https://www.adventureworks.com/Page2.aspx")

' Loop through and test your pages 
For Each Url In MyUrls
	GetWebPage Url
Next

Sub GetWebPage(ByVal Url)
	' You may have to change your version 
	Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
	oXMLHTTP.open "GET", Url
	oXMLHTTP.send
	' Test for response code 200
	If Not oXMLHTTP.status = 200 Then
		SendEmail	Url
	End If
End Sub

Sub SendEmail(Msg)
	' Set your SMTP server
	EmailSrv = "localhost"
	EmailFrom = "test@test.com"
	' I use my cell provider e-mail so i can get these e-mails on my PDA
	EmailTo = "help@test.com; mycellphone#@mycellprovider.com"
	Set objEmail = CreateObject("CDO.Message")
	objEmail.From = EmailFrom
	objEmail.To = EmailTo
	objEmail.Subject = "Web Monitor" 
	objEmail.Textbody = "Web-1 is not responding to HTTP requests. " & Msg
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = EmailSrv 
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objEmail.Configuration.Fields.Update
	objEmail.Send
End Sub

