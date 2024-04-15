'On Error Resume next

Dim TotalRam

ComputerName = InputBox ("Enter the computer name :" & chr(13) & chr(13) & _
"Leave it blank for local machine", "Computer Basic Info")

If ComputerName = Empty Then ComputerName = "."

Set Output = CreateObject("Scripting.FileSystemObject")
Set WshShell = Wscript.CreateObject("WScript.Shell")
Set WMIService = GetObject ("winmgmts:\\" & ComputerName & "\root\cimv2")
Set BIOSInfoList = WMIService.ExecQuery("Select * from Win32_BIOS",,48)
Set CompSysList = WMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)
Set ProcessorList = WMIService.ExecQuery("Select * from Win32_Processor",,48)
Set GETTIA = WMIService.ExecQuery("Select * from Win32_Environment Where Name = 'TIA'",,48)
Set HDList = WMIService.ExecQuery _
("Select * from Win32_LogicalDisk Where Description = 'Local Fixed Disk'",,48)
Set RAMList = WMIService.ExecQuery("Select * from Win32_PhysicalMemory",,48)
Set AppList = WMIService.ExecQuery("Select * from Win32Reg_AddRemovePrograms",,48)
Set NetList = WMIService.ExecQuery _
("Select * from Win32_NetworkAdapterConfiguration where IPEnabled = True",,48)

TotalRAM = 0

Set OutputFile = Output.CreateTextFile("C:\Temp\" & ComputerName & ".txt", true)
OutputFile.Writeline "This info is for " & ComputerName & " captured at " & Now
OutputFile.WriteBlankLines 1

For Each BIOSInfoItem in BIOSInfoList
	OutputFile.Writeline "Manufacturer		: " & BIOSInfoItem.Manufacturer
	OutputFile.Writeline "BIOS Name		: " & BIOSInfoItem.Name
	OutputFile.Writeline "Dell Service Tag	: " & BIOSInfoItem.SerialNumber
Next

For Each TIANumber in GETTIA
	OutputFile.WriteLine "TIA Number		: " & TIANumber.VariableValue
Next

For Each CompSysItem in CompSysList
	OutputFile.Writeline "Computer Model 		: " & CompSysItem.Model
	OutputFile.WriteLine "Logged-in User 		: " & CompSysItem.UserName
Next

For Each ProcessorItem in ProcessorList
	OutputFile.Writeline "Processor Name		: " & ProcessorItem.Name
Next

For Each RAMItem in RAMList
	TotalRAM = TotalRAM + RAMItem.Capacity
Next

OutputFile.WriteLine "RAM Capacity		: " & TotalRAM/1048576 & " MB"

For Each HDItem in HDList
	outputFile.Writeline "Total HD space		: " & HDItem.Size/1000000000 & " GB"
	OutputFile.Writeline "Available Free Space	: " & HDItem.FreeSpace/1000000000 & " GB"
Next

For Each NetItem in NetList
	strIPAddress = Join (NetItem.IPAddress,",")
	OutputFile.Writeline "IP Address		: " & strIPAddress
Next

OutputFile.WriteBlankLines 1

OutputFile.WriteLine "Add Remove Programs List :"

For Each AppItem in AppList
	OutputFile.Writeline AppItem.DisplayName & "					" & AppItem.Version
Next

WshShell.Run "Notepad " & "C:\Temp\" & Computername & ".txt", 1

Set ComputerName = Nothing
Set WMIService = Nothing
WScript.Quit

