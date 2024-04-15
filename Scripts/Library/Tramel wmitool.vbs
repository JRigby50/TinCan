'Written by Aaron Tramel
'November 5, 2010
strComputer = "."

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim objWMI : Set objWMI = GetObject("winmgmts:")
Dim colSettingsComp : Set colSettings = objWMI.ExecQuery("Select * from Win32_ComputerSystem")
Dim colSettingsBios : Set colSettingsBios = objWMI.ExecQuery("Select * from Win32_BIOS")
Dim objComputer, strModel, strSerial
For Each objComputer in colSettings 
strModel = objComputer.Model
Next
For Each objComputer in colSettingsBios 
strSerial = objComputer.SerialNumber
Next

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("c:\hardware.log", ForWriting, 1)
set osvc = getObject("winmgmts:root\cimv2")
Set devices = osvc.InstancesOf("Win32_PNPEntity Where ConfigManagerErrorCode > 0")
objTextFile.WriteLine(" ")
objTextFile.WriteLine("***************************************************")
objTextFile.WriteLine("DETECTED COMPUTER MODEL: " & strModel)
objTextFile.WriteLine("___________________________________________________")
objTextFile.WriteLine(" ")
objTextFile.WriteLine(" ")
objTextFile.WriteLine("PROBLEMS WERE FOUND WITH THE FOLLOWING DEVICES:")
For Each device In devices
objTextFile.WriteLine(device.Name & "," & device.PNPDeviceID)
Next
objTextFile.WriteLine("END OF PROBLEM DEVICE LIST")
objTextFile.Close


Set objTextFile = objFSO.OpenTextFile("c:\hardware.log", ForAppending, 1)

Set objWMIService = GetObject(_
    "winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_PnPEntity ")
objTextFile.WriteLine(" ")
objTextFile.WriteLine(" ")
objTextFile.WriteLine(" ")
objTextFile.WriteLine(" ")
objTextFile.WriteLine("ALL DEVICES DETECTED IN THE MACHINE*****************")
For Each objItem in colItems
objTextFile.WriteLine(objItem.Name)
objTextFile.WriteLine(objItem.PNPDeviceID)
objTextFile.WriteLine(" ")
Next
objTextFile.Close
Wscript.Echo "Device/PNP ID Written to c:\hardware.log for all devices"
