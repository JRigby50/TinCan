'****** re-used html needed later ******
'top listing
list = "<td><font size=" & chr(34) & "2" & chr(34) & _
	">Network Name</font></td><td><font size=" & chr(34) & "2" & chr(34) & _
        ">IP Address</font></td><td><font size=" & chr(34) & "2" & chr(34) & _
	">Model</font></td><td><font size=" & chr(34) & "2" & chr(34) & _
	">Comments</font></td><td><font size=" & chr(34) & "2" & chr(34) & _
	">Share Name</font></td></tr>"

'td open close
tco = "</font></td><td><font size=" & chr(34) & "2" & chr(34) & ">"

'open table
opentable = "<tr><td><font size=" & chr(34) & "2" & chr(34) & ">"

'close table
closetable = "</font></td></tr>"

'*************

strcomputer = "."

Set objFSO = CreateObject("Scripting.FileSystemObject")
'change C:\printers.htm to whatever you wish
Set File = objFSO.CreateTextFile("C:\printers.htm", True)


Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & _
    strComputer & "\root\cimv2")
Set colInstalledPrinters = objWMIService.ExecQuery ("Select * from Win32_Printer")

'-----------------------------------------------------------------
file.writeline "<html><head><title>Printer Inventory</title>"
file.writeline "</head>"
file.writeline "<body><table border=" & chr(34) & "0" & chr(34) & "width=" & chr(34) & _
    "1000" & chr(34) & "id=" & chr(34) & "table1" & chr(34) & ">" & list 

For Each objPrinter in colInstalledPrinters
	name = objprinter.name
	comment = objprinter.comment
	sharename = objprinter.sharename
	port = objprinter.portname
	drivername = objprinter.drivername
	
	file.writeline opentable & name & tco & port & tco & drivername & tco & _
    comment & tco & ShareName & closetable
Next

file.writeline "</font></table></body></html>"

File.Close

