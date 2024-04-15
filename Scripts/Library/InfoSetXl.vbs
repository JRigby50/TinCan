x = 2
Const ForReading = 1
Set objDictionary = CreateObject("Scripting.Dictionary")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    ("c:\serv3.txt" , forReading)
i = 0
Do Until objTextFile.AtEndOfStream
strNextLine = objTextFile.Readline
objDictionary.Add i, strNextLine
i = i + 1
loop

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add
Set objWorksheet = objWorkbook.Worksheets(1)

For Each objItem in objDictionary
strComputer = objDictionary.Item(objItem)

On Error Resume next
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 

objExcel.Cells(1, 1).Value = "Server name"
objExcel.Cells(1, 2).Value = "Manufacturer"
objExcel.Cells(1, 3).Value = "Model"
objExcel.Cells(1, 4).Value = "OS"
objExcel.Cells(1, 5).Value = "Service Pack"
objExcel.Cells(1, 6).Value = "Service Tag"

objExcel.Cells(x, 1).Value = strcomputer

Set ComputerSystem = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
For Each objComputerSystem in ComputerSystem

objExcel.Cells(x, 2).Value = objComputerSystem.Manufacturer
objExcel.Cells(x, 3).Value = objComputerSystem.Model

next
Set colOperatingSystems = objWMIService.ExecQuery _
("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems

objExcel.Cells(x, 4).Value = objOperatingSystem.Caption  
objExcel.Cells(x, 5).Value = objOperatingSystem.ServicePackMajorVersion
next
Set serial = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_BIOS",,48) 
For Each Item in serial

objExcel.Cells(x, 6).Value = Item.SerialNumber

x = x + 1

next
next
