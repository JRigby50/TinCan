'This code is simple eMail address validator that uses Regular Expression in VBScript.
'The code will accepts the email address using a Input box and passed to the test function of the Reg. Exp. and display the results in a message box.
'Any queries please forward to sskanagal@hotmail.com

'Create RE & other objects. 
Set objShell = CreateObject("WScript.Shell") 
Set objRegExp = CreateObject("vbscript.RegExp") 
 
objRegExp.Pattern = "^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$" 
eMailAddress = InputBox("Enter the full eMail address to validate:", "Enter email address") 
 
'Make it case sensitivity 
objRegExp.IgnoreCase = True 
 
'Execute the validator test 
returnVal = objRegExp.test(eMailAddress) 
If Not returnVal Then  
    Msgbox eMailAddress & " is not a Valid email address.", vbInformation+vbOkOnly, "Invalid address" 
Else 
    Msgbox eMailAddress & " is a valid email address.", vbInformation+vbOkOnly, "Valid address" 
End If 
