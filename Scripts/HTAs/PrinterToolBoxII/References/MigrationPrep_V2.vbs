'Script to create a text file Backup of the Print Server That can be edited and used with the SetupPrintServer.vbs script to
'setup both the Ports and Queues on a new Print Server.
'Updated 4/1/14 to use a better PING subroutine, to include the ShareName and to Replace Commas in strComment and strLocation with  " -"
'and to ignore the Microsoft XPS Document Writer queue
Option Explicit
'Define Variables and Constants
Dim strServer, strSNMPCommunity, strSNMPEnabled, strQuery, strPrinterName, strComment, strDriverName, strPortName, strLocation, strStatus, strIP, strShareName
Dim objWMIService, colInstalledPrinters, objPrinter, objPrintQueuesFile, colPorts, objPort, objWshShell, objFSO, colIPs, objIP, objPing, objStatus, bolResult, intLen
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const OverwriteExisting = True
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20
strServer = "." 
Set objWshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objPrintQueuesFile = objFSO.CreateTextFile(".\" & MachineName & "_printers.csv", OverwriteExisting)
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strServer & "\root\cimv2")
Set colInstalledPrinters = objWMIService.ExecQuery ("Select * from Win32_Printer")
'*****
objPrintQueuesFile.WriteLine "Printer Name,ShareName,Status,Port Name,Driver,Comment,Location,SNMP"
For Each objPrinter in colInstalledPrinters
	strPrinterName = objPrinter.Name
	strComment = objPrinter.Comment
	'Make sure it's not the XPS Queue
		If strPrinterName <> "Microsoft XPS Document Writer" Then
		'Replace commas
			If InStr(strComment, ",") Then
			strComment = Replace(strComment, ",", " -")
			End If
		strShareName = objPrinter.ShareName
		strPortName = objPrinter.PortName
		
		'Fix Formatting of PortName
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
		
		
		strDriverName = objPrinter.DriverName
		strLocation = objPrinter.Location
		'Replace commas
			If InStr(strLocation, ",") Then
			strLocation = Replace(strLocation, ", ", " - ")
			End If
		Ping
		SNMP
		objPrintQueuesFile.Writeline strPrinterName & "," & strShareName & "," & strStatus & "," & strPortName & "," & strDriverName & "," & strComment & "," & strLocation & "," & strSNMPCommunity
		End If
Next
objPrintQueuesFile.Close
MsgBox "Finished Collecting Migration Data."
Sub SNMP
strSNMPCommunity = ""
strQuery = "SELECT SNMPEnabled, SNMPCommunity FROM Win32_TCPIPPrinterPort Where Name = '" & strPortName & "'"
Set colPorts = objWMIService.ExecQuery(strQuery, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
	For Each objPort In colPorts
		strSNMPCommunity = objPort.SNMPCommunity
	Next
End Sub

Sub Ping
Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & strIP & "'")
	    For Each objStatus in objPing
	    bolResult = True
	        If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then 
	            bolResult = False
	        End If
		Next
        If bolResult = False Then 
            strStatus = "is NOT Pingable"
        Else
		    strStatus = "is Pingable"
        End If
End Sub
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