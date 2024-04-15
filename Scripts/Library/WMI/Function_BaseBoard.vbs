Option Explicit
Dim strServer, arrTest
strServer = "."

arrTest = BaseboardInfo(strServer)

Function BaseboardInfo(strComputer)
'Returns an Array of information about the computer baseboard.
Dim objWMIService, colItems, objItem, arrBoardInfo(), x
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
x = 0
Set colItems = objWMIService.ExecQuery("Select * from Win32_BaseBoard")
For Each objItem in colItems
	ReDim Preserve arrBoardInfo(23,x)
    arrBoardInfo(0,x) = "Depth: " & objItem.Depth
    arrBoardInfo(1,x) = "Description: " & objItem.Description
    arrBoardInfo(2,x) = "Height: " & objItem.Height
    arrBoardInfo(3,x) = "Hosting Board: " & objItem.HostingBoard
    arrBoardInfo(4,x) = "Hot Swappable: " & objItem.HotSwappable
    arrBoardInfo(5,x) = "Manufacturer: " & objItem.Manufacturer
    arrBoardInfo(6,x) = "Model: " & objItem.Model
    arrBoardInfo(7,x) = "Name: " & objItem.Name
    arrBoardInfo(8,x) = "Other Identifying Information: " & objItem.OtherIdentifyingInfo
    arrBoardInfo(9,x) = "Part Number: " & objItem.PartNumber
    arrBoardInfo(10,x) = "Powered On: " & objItem.PoweredOn
    arrBoardInfo(11,x) = "Product: " & objItem.Product
    arrBoardInfo(12,x) = "Removable: " & objItem.Removable
    arrBoardInfo(13,x) = "Replaceable: " & objItem.Replaceable
    arrBoardInfo(14,x) = "Requirements Description: " & objItem.RequirementsDescription
    arrBoardInfo(15,x) = "Requires DaughterBoard: " & objItem.RequiresDaughterBoard
    arrBoardInfo(16,x) = "Serial Number: " & objItem.SerialNumber
    arrBoardInfo(17,x) = "SKU: " & objItem.SKU
    arrBoardInfo(18,x) = "Slot Layout: " & objItem.SlotLayout
    arrBoardInfo(19,x) = "Special Requirements: " & objItem.SpecialRequirements
    arrBoardInfo(20,x) = "Tag: " & objItem.Tag
    arrBoardInfo(21,x) = "Version: " & objItem.Version
    arrBoardInfo(22,x) = "Weight: " & objItem.Weight
    arrBoardInfo(23,x) = "Width: " & objItem.Width
    x = x + 1
Next
BaseboardInfo = arrBoardInfo
End Function