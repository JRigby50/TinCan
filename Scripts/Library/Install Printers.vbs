'Create a FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Open a text file of Printers, Drivers, Ports, Location and ShareName with one Printer setup per line
Set objTS = objFSO.OpenTextFile("c:\printers.txt")
Set objTSOut = objFSO.CreateTextFile("c:\errors.txt")

'Go through the text file
Do Until objTS.AtEndOfStream

'Get next printer
  strPrinter = objTS.ReadLine
  myArray=Split(strPrinter,",")

 On Error Resume Next
  
'Install Printer
	 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objPrinter = objWMIService.Get("Win32_Printer").SpawnInstance_

objPrinter.DeviceID   = myArray(0)
objPrinter.DriverName = myArray(1)
objPrinter.PortName   = myArray(2)
objPrinter.Location = myArray(3)
objPrinter.Network = True
objPrinter.Shared = True
objPrinter.ShareName = myArray(0)
objPrinter.Put_	 

If Err <> 0 Then
 	objTSOut.WriteLine "Error creating  " & myArray(0) & ":" & Err.Number & ", " & Err.Description
 Else
	 On Error Goto 0
End If
 
Loop

objTS.Close
objTSOut.Close



 