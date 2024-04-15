Option Explicit
Dim strServer, arrTest
strServer = "."

arrTest = PPorts(strServer)

Function PPorts(strComputer)
'Returns information about the parallel ports installed on a computer
Dim objWMIService, colItems, objItem, arrPPorts(5,0), x
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ParallelPort",,48)
For Each objItem in colItems
    Wscript.Echo "Availability: " & objItem.Availability
    For Each strCapability in objItem.Capabilities
        Wscript.Echo "Capability: " & strCapability
    Next
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OS Auto Discovered: " & objItem.OSAutoDiscovered
    Wscript.Echo "PNP Device ID: " & objItem.PNPDeviceID
    Wscript.Echo "Protocol Supported: " & objItem.ProtocolSupported
Next
PPorts = arrPPorts
End Function