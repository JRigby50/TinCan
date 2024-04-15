Option Explicit
Dim strServer, arrTest
strServer = "."

arrTest = PhysMem(strServer)

Function PhysMem(strComputer)
' Returns an Array of information about the physical memory on a computer
Dim arrPhysMem(), objWMIService, colItems, objItem, x
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory",,48)
x = 0
For Each objItem in colItems
ReDim Preserve arrPhysMem(14,x)
    arrPhysMem(0,x) = "Bank Label: " & objItem.BankLabel
    arrPhysMem(1,x) = "Capacity: " & objItem.Capacity
    arrPhysMem(2,x) = "Data Width: " & objItem.DataWidth
    arrPhysMem(3,x) = "Description: " & objItem.Description
    arrPhysMem(4,x) = "Device Locator: " & objItem.DeviceLocator
    arrPhysMem(5,x) = "Form Factor: " & objItem.FormFactor
    arrPhysMem(6,x) = "Hot Swappable: " & objItem.HotSwappable
    arrPhysMem(7,x) = "Manufacturer: " & objItem.Manufacturer
    arrPhysMem(8,x) = "Memory Type: " & objItem.MemoryType
    arrPhysMem(9,x) = "Name: " & objItem.Name
    arrPhysMem(10,x) = "Part Number: " & objItem.PartNumber
    arrPhysMem(11,x) = "Position In Row: " & objItem.PositionInRow
    arrPhysMem(12,x) = "Speed: " & objItem.Speed
    arrPhysMem(13,x) = "Tag: " & objItem.Tag
    arrPhysMem(14,x) = "Type Detail: " & objItem.TypeDetail
    x = x + 1
Next
PhysMem = arrPhysMem
End Function