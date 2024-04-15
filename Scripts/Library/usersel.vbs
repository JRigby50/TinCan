function getInput( )
Dim answ 
timeOut = 30 
title = "Write Failure!" 
btype = 2
'create object 
Set w =WScript.CreateObject("WScript.Shell") 
getInput = w.Popup ("Error writing to the drive. Try again?",timeOut,title,btype)
End Function
answer = getInput( )
Select Case answer 
Case 3
WScript.Echo "You selected Abort." 
Case 4
WScript.Echo "You selected Retry." 
Case 5 
WScript.Echo "You selected Ignore." 
Case Else 
WScript.Echo "No selection in the time allowed. "
End Select 

