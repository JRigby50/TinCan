Option Explicit
Dim strServer, arrTest
strServer = "."

arrTest = SPorts(strServer)

Function SPorts(strComputer)
'Returns information about the serial ports installed on a computer
Dim objWMIService, colItems, objItem, arrSPorts(), x
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_SerialPort",,48)
x = 0 
For Each objItem in colItems
	ReDim Preserve arrSPorts(26,x)
    arrSPorts(0,x) =  "Binary: " & objItem.Binary
    arrSPorts(1,x) =  "Description: " & objItem.Description
    arrSPorts(2,x) =  "Device ID: " & objItem.DeviceID
    arrSPorts(3,x) =  "Maximum Baud Rate: " & objItem.MaxBaudRate
    arrSPorts(4,x) =  "Maximum Input Buffer Size: " & objItem.MaximumInputBufferSize
    arrSPorts(5,x) =  "Maximum Output Buffer Size: " & objItem.MaximumOutputBufferSize
    arrSPorts(6,x) =  "Name: " & objItem.Name
    arrSPorts(7,x) =  "OS Auto Discovered: " & objItem.OSAutoDiscovered
    arrSPorts(8,x) =  "PNP Device ID: " & objItem.PNPDeviceID
    arrSPorts(9,x) =  "Provider Type: " & objItem.ProviderType
    arrSPorts(10,x) =  "Settable Baud Rate: " & objItem.SettableBaudRate
    arrSPorts(11,x) =  "Settable Data Bits: " & objItem.SettableDataBits
    arrSPorts(12,x) =  "Settable Flow Control: " & objItem.SettableFlowControl
    arrSPorts(13,x) =  "Settable Parity: " & objItem.SettableParity
    arrSPorts(14,x) =  "Settable Parity Check: " & objItem.SettableParityCheck
    arrSPorts(15,x) =  "Settable RLSD: " & objItem.SettableRLSD
    arrSPorts(16,x) =  "Settable Stop Bits: " & objItem.SettableStopBits
    arrSPorts(17,x) =  "Supports 16-Bit Mode: " & objItem.Supports16BitMode
    arrSPorts(18,x) =  "Supports DTRDSR: " & objItem.SupportsDTRDSR
    arrSPorts(19,x) =  "Supports Elapsed Timeouts: " & objItem.SupportsElapsedTimeouts
    arrSPorts(20,x) =  "Supports Int Timeouts: " & objItem.SupportsIntTimeouts
    arrSPorts(21,x) =  "Supports Parity Check: " & objItem.SupportsParityCheck
    arrSPorts(22,x) =  "Supports RLSD: " & objItem.SupportsRLSD
    arrSPorts(23,x) =  "Supports RTSCTS: " & objItem.SupportsRTSCTS
    arrSPorts(24,x) =  "Supports Special Characters: " & objItem.SupportsSpecialCharacters
    arrSPorts(25,x) =  "Supports XOn XOff: " & objItem.SupportsXOnXOff
    arrSPorts(26,x) =  "Supports XOn XOff Setting: " & objItem.SupportsXOnXOffSet
    x= x + 1
Next
SPorts = arrSPorts
End Function