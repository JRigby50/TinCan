' FILE NAME: EventLog.vbs
' AUTHOR: Adam Hayes
' CREATION DATE : 12 Aug 2003
' REVISION DATE : 22 Jan 2004
' COMMENT: Collects all of the event log entries from any computer that
'          is listed in computers.txt. It creates an Excel file  
'    for each computer name and then clears all Event Logs.
' REVISED by Jim Rigby: 18 Aug 2008
' Removed AD query and Clearing of Event Logs

'==========================================================================
Option Explicit
On Error Resume Next
Dim strTxtPath,oFSO,strReportPath
Const ForReading = 1, ForWriting = 2
strReportPath="C:\LogReader"
strTxtPath="LogReader xt"
Set oFSO = CreateObject("Scripting.FileSystemObject")

'==========================================================================
' SUB NAME: MAIN
' AUTHOR: Adam Hayes
' CREATION DATE : 12 Aug 2003
' COMMENT:  Calls each sub procedure to run the script
'==========================================================================
Call MoveFiles()
Call GetLogNames()
Call GetEvents()
'Call ClearTxtFiles()
WScript.Echo "Event Log Script Complete"


'==========================================================================
' SUB NAME: GetLogNames()
' AUTHOR: Adam Hayes
' CREATION DATE : 14 August 2003
' REVISION DATE :
' COMMENT: Opens the computers.txt file and queries each computer for the
'    Event Logs that exist on each. It creates a txt file of the 
'    names and logs as logs.txt in the strTxtPath.
'==========================================================================
Sub GetLogNames ()
Dim oNameFile,strNextLine,arrComputerName,strComputerName,oLogFile
Dim objWMIService,colItems,objItem,strLogName,i

Set oNameFile = oFSO.OpenTextFile(strTxtPath&"computers.txt", ForReading)
Set oLogFile = oFSO.CreateTextFile(strTxtPath&"logs.txt", True)
Do Until oNameFile.AtEndOfStream
    strNextLine = oNameFile.Readline
    arrComputerName= Split(strNextLine,vbCRLF)
    For i=0 To UBound (arrComputerName)
 strComputerName=arrComputerName(i)
 oLogFile.Write(strComputerName)
 Set objWMIService = GetObject _
  ("winmgmts:" & strComputerName & "ootcimv2")
 Set colItems = objWMIService.ExecQuery _
  ("Select * from Win32_NTEventlogFile",,48)
 For Each objItem in colItems
     strLogName=objItem.logfilename
     oLogFile.Write(","& strLogName)
 Next
 oLogFile.Write (vbCRLF)
    Next
Loop
oLogFile.Close
oNameFile.Close
End Sub

'==========================================================================
' FILE NAME: MoveFiles()
' AUTHOR: Adam Hayes
' CREATION DATE : 14 August 2003
' REVISION DATE :
' COMMENT: Moves any existing XLS files from strReportPath into the "old"
'    subdirectory of strReportPath
'==========================================================================
Sub MoveFiles()
Dim oFolder,colFiles,intNumber
Set oFolder = oFSO.GetFolder(strReportPath)
Set colFiles = oFolder.Files
intNumber=colFiles.count
If intNumber0 Then
    oFSO.MoveFile strReportPath&"*.xls" , strReportPath&"old"
End If
End Sub

'==========================================================================
' SUB NAME: GetEvents()
' AUTHOR: Adam Hayes
' CREATION DATE : 14 August 2003
' REVISION DATE :
' COMMENT: Opens each ComputerNamelogs.txt from strTxtPath and queries 
'    each computer
'==========================================================================
Sub GetEvents()
On Error Resume Next
Dim oDate,oLogFile,strNextLine,arrLogFile,strComputerName,strLogName
Dim i,objWMIService,colLoggedEvents,oFile,objEvent,strMessage,arrMessage
Dim objExcel, objContainer, objChild

Const xlAscending = 1
Const xlDescending = 2
Const xlGuess = 0
Const xlTopToBottom = 1

oDate=GetDate()

Set oLogFile = oFSO.OpenTextFile(strTxtPath&"logs.txt", ForReading)
Do Until oLogFile.AtEndOfStream
    strNextLine = oLogFile.ReadLine
    Set objExcel = WScript.CreateObject("Excel.Application")
    objExcel.Visible = False
    objExcel.Workbooks.Add
    If strNextLine"" then
 arrLogFile= Split(strNextLine,",")
 strComputerName=arrLogFile(0)
 'WScript.Echo strComputerName
 For i = 1 to Ubound(arrLogFile)
     strLogName=arrLogFile(i)
     'WScript.Echo "   " & strLogName
     objExcel.ActiveWorkbook.Sheets.Add
     objExcel.ActiveSheet.Name = strLogName
     objExcel.ActiveSheet.Range("A1").Activate
     objExcel.ActiveCell.Value = "Record"
     objExcel.ActiveCell.Offset(0,1).Value = "Time"
     objExcel.ActiveCell.Offset(0,2).Value = "Event Code"
     objExcel.ActiveCell.Offset(0,3).Value = "Event Type"
     objExcel.ActiveCell.Offset(0,4).Value = "Source"
     objExcel.ActiveCell.Offset(0,5).Value = "Message"
     objExcel.ActiveSheet.Columns("A").ColumnWidth="6.14"
     objExcel.ActiveSheet.Columns("B").ColumnWidth="11.57"
     objExcel.ActiveSheet.Columns("C").ColumnWidth="9.71"
     objExcel.ActiveSheet.Columns("D").ColumnWidth="11.86"
     objExcel.ActiveSheet.Columns("E").ColumnWidth="40.00"
     objExcel.ActiveSheet.Columns("F").ColumnWidth="67.71"
     objExcel.Sheets(strLogName).Range("F1:F300").WrapText = True
     objExcel.ActiveCell.Offset(1,0).Activate 
     Set objWMIService = GetObject _
  ("winmgmts:"&"{impersonationLevel=impersonate}!" & _
  strComputerName &"ootcimv2")
     Set colLoggedEvents = objWMIService.ExecQuery _
  ("Select * from Win32_NTLogEvent Where Logfile ='" _
  & strLogName & "'")
     For Each objEvent in colLoggedEvents
  If objEvent.Message  ""Then
      strMessage=objEvent.Message
      arrMessage=Split(strMessage,vbCRLF)
      strMessage=Join(arrMessage)
      arrMessage=Split(strMessage,vbLF)
      strMessage=Join(arrMessage)
  End If
  objExcel.ActiveCell.Value = objEvent.RecordNumber
  objExcel.ActiveCell.Offset(0,1).Value = objEvent.TimeWritten
  objExcel.ActiveCell.Offset(0,2).Value = objEvent.EventCode
  objExcel.ActiveCell.Offset(0,3).Value = objEvent.Type
  objExcel.ActiveCell.Offset(0,4).Value = objEvent.SourceName
  objExcel.ActiveCell.Offset(0,5).Value = strMessage
  objExcel.ActiveCell.Offset(1,0).Activate 
     Next
     objExcel.ActiveCell.CurrentRegion.Select
     objExcel.Selection.Sort  _
  objExcel.Worksheets(strLogName).Range("D2"), _
  xlDescending, objExcel.Worksheets(1).Range("B2"), , _
  xlDescending, , , xlGuess,1,False,xlTopToBottom
 Next
 objExcel.Sheets("Sheet1").Delete
 objExcel.Sheets("Sheet2").Delete
 objExcel.Sheets("Sheet3").Delete
 objExcel.AlertBeforeOverwriting = False
 objExcel.ActiveWorkbook.SaveAs(strReportPath & oDate _
  & "_" & strComputerName & ".xls")
 objExcel.Quit
    End If
Loop
oLogFile.Close
End Sub

'==========================================================================
' SUB NAME: ClearTxtFiles
' AUTHOR: Adam Hayes
' CREATION DATE : 14 Aug 2003
' REVISION DATE :
' COMMENT:  Clears all txt files form the strTxtPath
'==========================================================================
Sub ClearTxtFiles()
Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")
oFSO.DeleteFile(strTxtPath&"*.txt")
End Sub

'==========================================================================
' FUNCTION NAME: GetDate
' AUTHOR: Adam Hayes
' CREATION DATE : 13 Aug 2003
' REVISION DATE :
' COMMENT: Uses the Date() command to retrieve the current date, then
'    reformats it into DayMonthYear format and returns the value
'==========================================================================
Function GetDate()
Dim oDate,arrDate,oMonth,oDay,oYear
oDate=Date()
arrDate=Split(oDate,"/")
oDay=arrDate(1)
oYear=arrDate(2)
Select Case arrDate(0)
    Case "1" 
 oMonth="Jan"
    Case "2"
 oMonth="Feb"
    Case "3"
 oMonth="Mar"
    Case "4"
 oMonth="Apr"
    Case "5"
 oMonth="May"
    Case "6"
 oMonth="June"
    Case "7"
 oMonth="July"
    Case "8"
 oMonth="Aug"
    Case "9"
 oMonth="Sept"
    Case "10"
 oMonth="Oct"
    Case "11"
 oMonth="Nov"
    Case "12"
 oMonth="Dec"
End Select
arrDate(0)=oDay
arrDate(1)=oMonth
arrDate(2)=oYear
GetDate= Join(arrDate,"")
End Function
Return
             


 