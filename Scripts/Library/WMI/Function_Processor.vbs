Option Explicit
Dim strServer, arrTest
strServer = "RSEMD4-BACKUP01"

arrTest = ProcessorInfo(strServer)
MsgBox "OK?"
Function ProcessorInfo(strComputer)
'Returns an Array of information about the processors.
Dim objWMIService, colItems, objItem, arrPInfo(), x
x = 0
On Error Resume Next
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
For Each objItem in colItems
    ReDim Preserve arrPInfo(28,x)
    arrPInfo(0,x) = "Address Width: " & objItem.AddressWidth
    arrPInfo(1,x) = "Architecture: " & objItem.Architecture
    arrPInfo(2,x) = "Availability: " & objItem.Availability
    arrPInfo(3,x) = "CPU Status: " & objItem.CpuStatus
    arrPInfo(4,x) = "Current Clock Speed: " & objItem.CurrentClockSpeed
    arrPInfo(5,x) = "Data Width: " & objItem.DataWidth
    arrPInfo(6,x) = "Description: " & objItem.Description
    arrPInfo(7,x) = "Device ID: " & objItem.DeviceID
    arrPInfo(8,x) = "Ext Clock: " & objItem.ExtClock
    arrPInfo(9,x) = "Family: " & objItem.Family
    arrPInfo(10,x) = "L2 Cache Size: " & objItem.L2CacheSize
    arrPInfo(11,x) = "L2 Cache Speed: " & objItem.L2CacheSpeed
    arrPInfo(12,x) = "Level: " & objItem.Level
    arrPInfo(13,x) = "Load Percentage: " & objItem.LoadPercentage
    arrPInfo(14,x) = "Manufacturer: " & objItem.Manufacturer
    arrPInfo(15,x) = "Maximum Clock Speed: " & objItem.MaxClockSpeed
    arrPInfo(16,x) = "Name: " & objItem.Name
    arrPInfo(17,x) = "PNP Device ID: " & objItem.PNPDeviceID
    arrPInfo(18,x) = "Processor Id: " & objItem.ProcessorId
    arrPInfo(19,x) = "Processor Type: " & objItem.ProcessorType
    arrPInfo(20,x) = "Revision: " & objItem.Revision
    arrPInfo(21,x) = "Role: " & objItem.Role
    arrPInfo(22,x) = "Socket Designation: " & objItem.SocketDesignation
    arrPInfo(23,x) = "Status Information: " & objItem.StatusInfo
    arrPInfo(24,x) = "Stepping: " & objItem.Stepping
    arrPInfo(25,x) = "Unique Id: " & objItem.UniqueId
    arrPInfo(26,x) = "Upgrade Method: " & objItem.UpgradeMethod
    arrPInfo(27,x) = "Version: " & objItem.Version
    arrPInfo(28,x) = "Voltage Caps: " & objItem.VoltageCaps
    x = x + 1
Next
ProcessorInfo = arrPInfo
End Function