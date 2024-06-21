Option Explicit
Dim strServer
strServer = Wscript.Arguments.Item(0)
'******HTML strings used later ********************************************************************
'Toggle Table
Dim strToggle, strOpenTable, strCloseTable, strHeaderDrivers
Dim objFSO, objHTMLFile, objTextFile, objWMIService, colInstalledDrivers, objDriver
strToggle = "</font></td><td><font size=" & chr(34) & "1" & chr(34) & ">"
'Open Table
strOpenTable = "<tr><td><font size=" & chr(34) & "1" & chr(34) & ">"
'Close Table
strCloseTable = "</font></td></tr>"
'Header
strHeaderDrivers = "<th><font color=blue font size=" & chr(34) & "2" & chr(34) & _
	">Name</font></th><th><font color=blue font size=" & chr(34) & "2" & chr(34) & _
	">DriverVersion</font></th><th><font color=blue font size=" & chr(34) & "2" & chr(34) & _
	">Path</font></th><th><font color=blue font size=" & chr(34) & "2" & chr(34) & _
	">MonitorName</font></th></tr>"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objHTMLFile = objFSO.CreateTextFile("display.htm", True)
Set objTextFile = objFSO.CreateTextFile("drivers.txt", True)
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strServer & "\root\cimv2")
Set colInstalledDrivers =  objWMIService.ExecQuery ("Select * from Win32_PrinterDriver")
objHTMLFile.writeline "<html><head><title>Print Drivers</title>"
objHTMLFile.writeline "</head>"
objHTMLFile.writeline "<body><table border=" & chr(34) & "1" & chr(34) & "width=" & chr(34) & _
    "900" & chr(34) & "id=" & chr(34) & "table1" & chr(34) & ">" & strHeaderDrivers
		For Each objDriver in colInstalledDrivers
			objHTMLFile.writeline strOpenTable & objDriver.Name & strToggle & CStr(objDriver.Version) & strToggle & _
		    objDriver.FilePath & strToggle & objDriver.MonitorName & strCloseTable
		    objTextFile.WriteLine objDriver.Name
		Next
			If Err <> 0 And Err <> 451 And Err <> 454 And Err <> -2147023174 Then
			objHTMLFile.WriteLine "Sorry but an error occured with server " & strServer & " Error: " & Err
			End If
objHTMLFile.writeline "</font></table></body></html>"
objHTMLFile.Close
objTextFile.Close
objHeaderFile.Close