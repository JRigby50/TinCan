Description: Retrieves available drive space information from a set of computers and -- if available space is less than 80 gigabytes -- writes the information to an Excel spreadsheet. As written, the list of computers (Servers.txt) must be in the C:\ folder and must contain the name of one server per line. The Excel spreadsheet is created and data is added, but is not saved. 

On Error Resume Next

Const xpRangeAutoFormatList2 = 11
Const xlDescending = 2

'folderinfo represents the share you want to search!

'The following Excel VBA code opens a Excel Worksheet and sets up the columns for Vbscript data

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)
objExcel.Cells(1, 1).Value = "Server"
objExcel.Cells(1, 1).Font.Italic = True
objExcel.Cells(1, 1).Font.Bold = True
objExcel.Cells(1, 2).Value = "Drive Letter"
objExcel.Cells(1, 2).Font.Italic = True
objExcel.Cells(1, 2).Font.Bold = True
objExcel.Cells(1, 3).Value = "Drive Size"
objExcel.Cells(1, 3).Font.Italic = True
objExcel.Cells(1, 3).Font.Bold = True
objExcel.Cells(1, 4).Value = "Free Space"
objExcel.Cells(1, 4).Font.Italic = True
objExcel.Cells(1, 4).Font.Bold = True
objExcel.Cells(1, 5).Value = "Free Space %"
objExcel.Cells(1, 5).Font.Italic = True
objExcel.Cells(1, 5).Font.Bold = True

x = 2

'The following Vbscript firsts echos message prompt to the screen. Then runs Vbscript code
'to locate the Drive Infromation which is then imported into the opened Excel Spreadsheet.

WScript.Echo
WScript.Echo
WScript.Echo "Process started at: " & Now
WScript.Echo ""
WScript.Echo ""
WScript.Echo "Please wait. While completing your Drive Space Spreadsheet ..................."
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("c:\scripts\servers.txt", 1)

Do Until objFile.AtEndOfStream
  strComputer = objFile.ReadLine
  Set objWMIService = GetObject("winmgmts://" & strComputer)

  Set colDisks = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_LogicalDisk WHERE DriveType = 3")
  For Each objDisk In colDisks
    If objDisk.Freespace < 80000000000 Then     
      objExcel.Cells(x, 1) = strComputer 
      objExcel.Cells(x, 2)= objDisk.DeviceID 
      objExcel.Cells(x, 3)= FormatNumber(objDisk.Size/1024,0) 
      objExcel.Cells(x, 4)= FormatNumber(objDisk.FreeSpace/1024,0)  
      objExcel.Cells(x, 5)= FormatPercent(objDisk.FreeSpace/objDisk.Size, 0) 
      x = x + 1
    End If
     
  Next
Loop
Set objRange = objWorksheet.UsedRange
objRange.AutoFormat(xpRangeAutoFormatList2)
objRange.EntireColumn.Autofit()
Set objRange5 = objExcel.Range("E1")
objRange.Sort objRange5,,,,,,,1

WScript.Echo 
WScript.Echo
WScript.Echo "Your query completed at: " & Now

