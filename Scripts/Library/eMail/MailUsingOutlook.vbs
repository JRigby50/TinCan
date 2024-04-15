'========================================================================== 
' 
' NAME: MailUsingOutlook.vbs 
' 
' COMMENT: This script generates an e-mail using the Outlook client. 
' 
'========================================================================== 
 
'Create an Outlook object 
Dim Outlook 'As New Outlook.Application 
Set Outlook = CreateObject("Outlook.Application") 
   
'Create e new message 
Dim Message 'As Outlook.MailItem 
Set Message = Outlook.CreateItem(olMailItem) 
With Message 
    .Subject = "Important message from a script!" 
    .Body = "This message was generated from a script using the Outlook client." 
 
    'Set destination email address 
    .Recipients.Add ("k.myer@fabrikam.com") 
 
    'Set sender address. 
    Const olOriginator = 0 
    .Recipients.Add("sysadmin@fabrikam.com").Type = olOriginator 
    .Recipients.ResolveAll 
 
    'Send the Message 
    .Send 
End With 
