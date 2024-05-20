'Script to create a text file of the Print Online/Offline Status.
'Written by Jim Rigby james.rigby@ngc.com
'Last updated 3/28/14 by Jim Rigby
Option Explicit
'Define Variables and Constants
Dim strOSVer, strComputer, strServer, dtmThisDay, dtmThisMonth, dtmThisYear, objWshShell, bolSuccess, objFSO, strCMD, objPrintPortFile
Dim strPrinterName, strPortName, strDriverName, colInstalledPrinters, objPrinter, objWMIService, bolResult, objPing, strIP, intLen, objStatus
Dim strComment
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const OverwriteExisting = True
strComputer = "." 
dtmThisDay = Day(Date)
dtmThisMonth = Month(Date)
dtmThisYear = Year(Date)
Set objWshShell = CreateObject("WScript.Shell")
	'*****Generate .txt file*****
	strcomputer = "."
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objPrintPortFile = objFSO.CreateTextFile(MachineName & "_PrintersOffline.csv", OverwriteExisting)
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colInstalledPrinters = objWMIService.ExecQuery ("Select * from Win32_Printer")
	'*****
	objPrintPortFile.WriteLine "Printer Name,Port Name"
	On Error Resume Next
	For Each objPrinter in colInstalledPrinters
		strPrinterName = objPrinter.Name
		strPortName = objPrinter.PortName
		strDriverName = objPrinter.DriverName
		If InStr(strPortName,"IP_")= 1 Then
			intLen = Len (strPortName) - 3
			strIP = Right (strPortName,intLen)
			Else
			strIP = strPortName
		End If
		
		If InStr(strIP,":")<> 0 Then
			intLen = Len (strIP) - 3
			strIP = Left (strIP,(InStr(strIP,":") - 1))
		End If
		
'MsgBox "IP " & strIP & "|Port Name " & strPortName
    Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & strIP & "'")
	    For Each objStatus in objPing
	    bolResult = True
	        If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then 
	            bolResult = False
	        End If
		Next
		If bolResult = False Then
			'Make sure it's not the XPS Queue
			If strPrinterName <> "Microsoft XPS Document Writer" AND InStr(strPortName,"LPT") = 0 And InStr(strDriverName,"Designjet") = 0 Then
			objPrintPortFile.Writeline strPrinterName & "," & strPortName
			'Update the Queue
			strComment = objPrinter.Comment
			strComment = "(Note: No PING " & Now & ") " & strComment
			objPrinter.Comment = strComment
			objPrinter.Put_
			End If
		End If
	Next
	objPrintPortFile.Close
MsgBox "Finished Checking for Offline Printers", vbInformation, "Ping Printer Ports"

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