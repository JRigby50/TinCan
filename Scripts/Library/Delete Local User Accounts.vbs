'************************************************* 
' File:			Delete Local User Accounts.vbs 
' Author:		Jim Rigby 
' Version:     1.0 
' 
'************************************************* 
 
Set objShell = CreateObject("Wscript.Shell") 
Set objNetwork = CreateObject("Wscript.Network") 
 
strComputer = objNetwork.ComputerName 
 
Set colAccounts = GetObject("WinNT://" & strComputer & "") 
 
colAccounts.Filter = Array("user") 
    Message = Message & "Local User accounts:" & vbCrLf & vbCrLf 
 
For Each objUser In colAccounts 
 
    If objUser.Name <> "Administrator" AND objUser.Name <> "ASPNET" Then 
            Message = Message & objUser.Name 
            If objUser.AccountDisabled = TRUE then 
                 Message = Message & " is currently disabled" & vbCrLf 
            Else 
                Message = Message & " was enabled" & vbCrLf 
                objUser.AccountDisabled = True 
                objUser.SetInfo 
            End if 
    End If 
 
Next 
 
' Initialize title text. 
Title = "Local User Accounts By Andrew Barnes" 
objShell.Popup Message, , Title, vbInformation + vbOKOnly