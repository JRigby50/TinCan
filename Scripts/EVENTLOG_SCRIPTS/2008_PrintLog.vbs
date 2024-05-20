'Script to extract who is printing to what printer(s)on a Windows Server 2008 machine
'Written by Jim Rigby james.rigby@ngc.com 5/23/18
Option Explicit
Dim objFSO, objPrintLogFile, objWMIService, colRetrievedEvents, objEvent, objPrintLog
Dim strServer, strQuery, strDate, strYear, strMonth, strDay, strHour, strMinute, strTime, strMessage, strUser, strDomain, strEntry

Dim intTest

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const OverwriteExisting = True
Set objFSO = WScript.CreateObject("Scripting.Filesystemobject")
Set objPrintLogFile = objFSO.CreateTextFile(".\" & MachineName & "_PrintLog.csv", OverwriteExisting)
objPrintLogFile.WriteLine  "Domain,User,Date,Time,Message"
strServer = "."
'%SystemRoot%\System32\Winevt\Logs\Microsoft-Windows-PrintService%4Operational.evtx
strQuery = "Select * from Win32_NTLogEvent Where Logfile = 'Microsoft-Windows-PrintService%4Operational' and SourceName = 'PrintService'"' and Type = 'Information' and EventID = '307'"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strServer & "\root\cimv2")
Set colRetrievedEvents = objWMIService.ExecQuery (strQuery)
On Error Resume Next
For Each objEvent in colRetrievedEvents
'Parse the data and generate a cleaned up list.
	'Format Domain and User Name
	strUser = objEvent.User
	intTest = InStr(strUser, "\")
	strDomain = Left(strUser,(InStr(strUser, "\") -1))
	strUser = Right(strUser, (Len(strUser) - (InStr(strUser, "\"))))
	'Format the date and time
	strDate = objEvent.TimeWritten
	strYear = Left(strDate,4)
	strMonth = Mid(strDate,5,2)
	strDay = Mid(strDate,7,2)
	strHour = Mid(strDate,9,2)
	strMinute = Mid(strDate,11,2)
	strDate = strMonth & "/" & strDay & "/" & strYear
	strTime = strHour & ":" & strMinute
	'Clean the Message String
	strMessage = objEvent.General
	strMessage = Right(strMessage, (Len(strMessage) - (InStr(strMessage, "was printed on") + 14)))
	strMessage = Left(strMessage, (InStr(strMessage, " through port")))
	
	WScript.Echo strDomain
	WScript.Echo strUser
	WScript.Echo strMessage
	
	'Write the data to the PrintLog
'	objPrintLogFile.Write strDomain & ","
'	objPrintLogFile.Write strUser & ","
'	objPrintLogFile.Write strDate & ","
'	objPrintLogFile.Write strTime & ","
'	objPrintLogFile.WriteLine strMessage
Next
objPrintLogFile.Close
'Remove Duplicates
Set objPrintLog = objFSO.OpenTextFile(".\" & MachineName & "_PrintLog.csv", ForReading)
Do Until objPrintLog.AtEndOfStream
	strEntry = objPrintLog.ReadLine
	
Loop
objPrintLog.Close
MsgBox "The script has completed collecting the Print Log"
'*****Subroutines And Functions*****
Function MachineName
Dim objWMIService, colItems, objItem, strComputer, strName
strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48) 
For Each objItem in colItems 
    strName = objItem.CSName
Next
MachineName = strName
End Function