Option Explicit

Function CreateHTMLMsg(fileHTML)' As String /As Outlook.MailItem
Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem
Dim objFSO As Scripting.FileSystemObject
Dim objStream As Scripting.TextStream
Dim strHTMLFile As String
	    On Error Resume Next
	    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(fileHTML) Then
        Set objOL = Application
        Set objMsg = objOL.CreateItem(olMailItem)
        Set objStream = objFSO.OpenTextFile(fileHTML, ForReading)
        objMsg.HTMLBody = objStream.ReadAll
    End If
Set CreateHTMLMsg = objMsg
Set objOL = Nothing
Set objMsg = Nothing
Set objFSO = Nothing
Set objStream = Nothing
End Function