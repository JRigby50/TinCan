on Error Resume Next

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add

objExcel.Cells(1, 1).Value = "  Computer Name  "
objExcel.Cells(1, 1).Font.Bold = True
objExcel.Cells(1, 1).Interior.ColorIndex = 45
objExcel.Cells(1, 1).Alignment = "Center"
Set objRange = objExcel.ActiveCell.EntireColumn
objRange.AutoFit()

objExcel.Cells(1, 2).Value = "  Manufacturer  "
objExcel.Cells(1, 2).Font.Bold = True
objExcel.Cells(1, 2).Interior.ColorIndex = 45
Set objRange = objExcel.Range("B1")
objRange.Activate
Set objRange = objExcel.ActiveCell.EntireColumn
objRange.AutoFit()

objExcel.Cells(1, 3).Value = "  Model  "
objExcel.Cells(1, 3).Font.Bold = True
objExcel.Cells(1, 3).Interior.ColorIndex = 45
Set objRange = objExcel.Range("C1")
objRange.Activate
Set objRange = ObjExcel.ActiveCell.EntireColumn
objRange.AutoFit()

objExcel.Cells(1, 4).Value = "  RAM  "
objExcel.Cells(1, 4).Font.Bold = True
objExcel.Cells(1, 4).Interior.ColorIndex = 45
Set objRange = objExcel.Range("C1")
objRange.Activate
Set objRange = ObjExcel.ActiveCell.EntireColumn
objRange.AutoFit()

y = 2
Set fso1 = CreateObject("Scripting.FileSystemObject")
Set pcfile = fso1.OpenTextFile("c:\windowsscripts\pc-names.txt",1)
do while Not pcfile.AtEndOfStream
    computerName = pcfile.readline
Err.Clear
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
    computerName & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
If Err.Number = 0 then
    For Each objComputer in colSettings
        strManufacturer = objComputer.Manufacturer
        strModel = objComputer.Model
        strRAM = FormatNumber(((objComputer.TotalPhysicalMemory / 1024) / 1024), 2)
    
        if computerName <> oldName then
          objExcel.Cells(y, 1).Value = computerName
          objExcel.Cells(y, 1).Alignment = "Center"
          objExcel.Cells(y, 2).Value = strManufacturer
          objExcel.Cells(y, 2).Alignment = "Center"
          objExcel.Cells(y, 3).Value = strModel
          objExcel.Cells(y, 3).Alignment = "Center" 
          objExcel.Cells(y, 4).Value = strRAM
          objExcel.Cells(y, 4).Alignment = "Center"
          oldName = computerName
          y = y + 1
        end if
       computerName = ""
       strManufacturer = ""
       strModel = ""
       strRAM = ""
       Err.clear
    next
Else
         objExcel.Cells(y, 1).Value = computerName
         objExcel.Cells(y, 1).Alignment = "Center"
         objExcel.Cells(y, 2).Value = "Not on line"
         objExcel.Cells(y, 2).Alignment = "Center"
         y = y + 1
         Err.clear 
End If
Loop
y = y + 1
objExcel.Cells(y, 1).Value = "Scan Complete"
objExcel.Cells(y, 1).Font.Bold = True

et objRange = objExcel.Range("A1")
objRange.Activate
Set objRange = objExcel.ActiveCell.EntireColumn
objRange.Autofit()
Set objRange = objExcel.Range("B1")
objRange.Activate
Set objRange = objExcel.ActiveCell.EntireColumn
objRange.Autofit() 
Set objRange = objExcel.Range("C1")
objRange.Activate
Set objRange = objExcel.ActiveCell.EntireColumn
objRange.Autofit()
Set objRange = objExcel.Range("D1")
objRange.Activate
Set objRange = objExcel.ActiveCell.EntireColumn
objRange.Autofit()

