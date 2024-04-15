'************************************************* 
' Remove User Accounts From Local Group.vbs 
' Jim Rigby 
' V 1.0 
' June 28,2011
'************************************************* 
 
' get target machine name and password 
' from command line arguments 
strServer = wscript.arguments(0) 
strPwd = wscript.arguments(1) 
strUser = "ITAdmin" 
 
' connect to target machine and  
' create new user 
Set oServer = GetObject ("WinNT://" & strServer) 
Set oUser = oServer.Create ("user", strUser) 
oUser.SetPassword strPwd 
oUser.SetInfo 
 
' Remove User from 'Administrators' group 
Set objGroup = GetObject("WinNT://" & strServer & "/administrators,group") 
objGroup.Remove(oUser.ADsPath) 
Group.Setinfo 
 
' release objects 
Set oServer = nothing 
Set oUser = nothing 
Set group= nothing 
