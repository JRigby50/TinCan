strToggle = "</font></td><td><font size=" & chr(34) & "2" & chr(34) & ">"
strToggleBlack = "</font></td><td><font size=" & chr(34) & "2" & chr(34) & ">"
strToggleRed = "</font></td><td><font color=red font size=" & chr(34) & "2" & chr(34) & ">"
'Open Table
strOpenTable = "<tr><td><font size=" & chr(34) & "2" & chr(34) & ">"
'Open Table, TAB
strOpenTableTab = "<tr><td><font size=" & chr(34) & "3" & chr(34) & ">"
'Close Table
strCloseTable = "</font></td></tr>"


strHeaderDrivers = "<th><font color=blue font size=" & chr(34) & "2" & chr(34) & _
	">Name</font></th><th><font color=blue font size=" & chr(34) & "2" & chr(34) & _
	">Version</font></th><th><font color=blue font size=" & chr(34) & "2" & chr(34) & _
    ">DriverVersion</font></th><th><font color=blue font size=" & chr(34) & "2" & chr(34) & _
	">Path</font></th><th><font color=blue font size=" & chr(34) & "2" & chr(34) & _
	">Environment</font></th><th><font color=blue font size=" & chr(34) & "2" & chr(34) & _
	">DriverEnv</font></th><th><font color=blue font size=" & chr(34) & "2" & chr(34) & _
	">Environment</font></th><th><font color=blue font size=" & chr(34) & "2" & chr(34) & _
	">MonitorName</font></th></tr>"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile("C:\Documents and Settings\rigbyja\Desktop\Scripts\HTAs\PrinterMasterBlaster\html\display.htm", True)
set objMaster = CreateObject("PrintMaster.PrintMaster.1")
objTextFile.writeline "<html><head><title>Print Drivers</title>"
objTextFile.writeline "</head>"
objTextFile.writeline "<body><table border=" & chr(34) & "1" & chr(34) & "width=" & chr(34) & _
    "1000" & chr(34) & "id=" & chr(34) & "table1" & chr(34) & ">"

Set objHeaderFile = objFSO.CreateTextFile("C:\Documents and Settings\rigbyja\Desktop\Scripts\HTAs\PrinterMasterBlaster\html\header.htm", True)  
objHeaderFile.writeline "<html><head><title>Print Driver Header</title>"
objHeaderFile.writeline "</head>"
objHeaderFile.writeline "<body><table border=" & chr(34) & "1" & chr(34) & "width=" & chr(34) & _
    "1000" & chr(34) & "id=" & chr(34) & "table1" & chr(34) & ">" & strHeaderDrivers
objHeaderFile.Close

dim objMaster

'The following code creates the required PrintMaster and Driver objects.
set objMaster = CreateObject("PrintMaster.PrintMaster.1")

strServer = ""
On Error Resume Next

For Each objDriver in objMaster.Drivers(strServer)
	objTextFile.writeline strOpenTable & objDriver.ModelName & strToggle & objDriver.Version & strToggle & CStr(objDriver.DriverVersion) & strToggle & _
    objDriver.Path & strToggle & objDriver.Environment & strToggle & objDriver.DriverArchitecture & strToggle & objDriver.MonitorName & strCloseTable
Next
	If Err <> 0 And Err <> 451 And Err <> 454 And Err <> -2147023174 Then
	objErrorsOut.WriteLine "Sorry but an error occured with server " & strServer & " Error: " & Err
	'MsgBox "Sorry but an error occured with server " & strServer & " Error: " & Err
	End If
objTextFile.writeline "</font></table></body></html>"
objTextFile.Close
objErrorsOut.Close
'Next