Option Explicit
Dim strServer, arrTest
strServer = "."

arrTest = BusInfo(strServer)
MsgBox "OK?"

Function BusInfo(strComputer)
' Returns information about the computer bus.
Dim x, objWMIService, colItems, objItem, arrBusInfo()
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Bus")
x=0
For Each objItem in colItems
	ReDim Preserve arrBusInfo(5,x)
    arrBusInfo(0,x) = "Bus Number: " & objItem.BusNum
    arrBusInfo(1,x) = "Bus Type: " & objItem.BusType
    arrBusInfo(2,x) = "Description: " & objItem.Description
    arrBusInfo(3,x) = "Device ID: " & objItem.DeviceID
    arrBusInfo(4,x) = "Name: " & objItem.Name
    arrBusInfo(5,x) = "PNP Device ID: " & objItem.PNPDeviceID
    x=x+1
Next
BusInfo = arrBusInfo
End Function