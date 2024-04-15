Option Explicit
Dim strServer, arrTest
strServer = "."
arrTest = IRQ(strServer)
MsgBox "OK?"

Function IRQ(strComputer)
'Returns an Array of information about the IRQ settings on a computer.
Dim objWMIService, colItems, objItem, arrIRQ(), x
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_IRQResource")
x = 0
For Each objItem in colItems
ReDim Preserve arrIRQ(5,x)
    arrIRQ(0,x) = "Availability: " & objItem.Availability
    arrIRQ(1,x) = "Hardware: " & objItem.Hardware
    arrIRQ(2,x) = "IRQ Number: " & objItem.IRQNumber
    arrIRQ(3,x) = "Name: " & objItem.Name
    arrIRQ(4,x) = "Trigger Level: " & objItem.TriggerLevel
    arrIRQ(5,x) = "Trigger Type: " & objItem.TriggerType
    x=x+1
Next
IRQ=arrIRQ
End Function