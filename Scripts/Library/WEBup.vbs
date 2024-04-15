'****************************************************
' Website Monitoring Script By Manupriya Jayathilake
' 23 January 2006
'****************************************************


url="http://www.microsoft.com"

on error resume next

Set objHTTP = CreateObject("MSXML2.XMLHTTP")
Call objHTTP.Open("GET", url, FALSE)

stat=objHTTP.status

if not stat="200" then
Set objEmail = CreateObject("CDO.Message")
objEmail.From = "manupriya_j@hotmail.com"
objEmail.To = "manupriya_j@hotmail.com" 
objEmail.Subject = "Your Website Down " 
objEmail.Textbody = url & "replied with status code " & stat & vbcrlf & _
    " Please Check and Restore the Web Service ASAP." & vbcrlf & now ()
objEmail.Send

end if

