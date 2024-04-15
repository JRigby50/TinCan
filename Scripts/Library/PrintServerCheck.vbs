Option Explicit

Set objFSO = CreateObject("Scripting.FileSystemObject")
strComputer = objNetwork.ComputerName
Set objTextFile = objFSO.CreateTextFile("c:\" & strcomputer & " Installed Fonts.txt", True)
objTextFile.WriteLine "Installed Fonts"
objTextFile.WriteLine

Installed_Fonts

Sub Installed_Fonts
'Checks Installed Fonts
Const FONTS = &H14&
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(FONTS)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path
Set colItems = objFolder.Items
For Each objItem in colItems
    Wscript.Echo objItem.Name
Next
End Sub
objTextFile.Close

'Creates Report of Basic Machine Info
Sub Basic_Report
 	Set objNetwork = CreateObject("Wscript.Network")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strComputer = objNetwork.ComputerName
	Set objTextFile = objFSO.CreateTextFile("c:\" & strcomputer & " Basic Info.txt", True)
On Error Resume Next
		objTextFile.WriteLine "Basic Info"
		objTextFile.WriteLine

	Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strcomputer & "\root\cimv2")
If Err.Number <> 0 Then
    MsgBox Err.Description
    Err.Clear
Else
	Set colBIOS = objWMIService.ExecQuery ("SELECT * FROM Win32_BIOS")
For Each objBIOS in colBIOS
	     objTextFile.WriteLine"Computer name: " & (strcomputer)
         objTextFile.WriteLine"Manufacturer: " & (objBIOS.Manufacturer)
         objTextFile.WriteLine"Serial number: " & (objBIOS.SerialNumber)
         objTextFile.WriteLine"BIOS Version: " & (objBIOS.SMBIOSBIOSVersion)
Next
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem",,48) 
For Each objItem in colItems 
	
	    objTextFile.WriteLine "Domain: " & (objItem.Domain)
	    'objTextFile.WriteLine "Manufacturer: " & (objItem.Manufacturer)
	    objTextFile.WriteLine "Model: " & (objItem.Model)
	    'objTextFile.WriteLine "Name: " & (objItem.Name)
	    objTextFile.WriteLine "NumberOfProcessors: " & (objItem.NumberOfProcessors)
	    objTextFile.WriteLine "TotalPhysicalMemory: " & (objItem.TotalPhysicalMemory)
Next
End If
objTextFile.Close
End Sub

Sub OS_Report
Dim objFSO, strComputer, objWMIService, colItems, objItem
	Set objNetwork = CreateObject("Wscript.Network")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strComputer = objNetwork.ComputerName
	Set objTextFile = objFSO.CreateTextFile("c:\" & strcomputer & " OS.txt", True)
	On error resume Next
		objTextFile.WriteLine "OS Version"
		objTextFile.WriteLine
		Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery("SELECT * " & "FROM Win32_OperatingSystem")
For Each objItem In colItems		
		 objTextFile.WriteLine "OS Info ror " & objItem.csname
		 objTextFile.WriteLine"OS: " & objItem.Caption
		 objTextFile.WriteLine"Build= " & objItem.BuildNumber 'objItem.BuildType
		 objTextFile.WriteLine"Service Pack= " & objItem.CSDVersion
		 objTextFile.WriteLine"Free Memory= " & objItem.FreePhysicalMemory    
		 objTextFile.WriteLine"Windows Directory= " & objItem.WindowsDirectory
Next
objTextFile.Close
End Sub

'Creates Report of Installed Patches
Sub Patches_Report
    On error resume Next
	Set objNetwork = CreateObject("Wscript.Network")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strComputer = objNetwork.ComputerName
	Set objTextFile = objFSO.CreateTextFile("c:\" & strcomputer & " Patches.csv", True)
	Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colQuickFixes = objWMIService.ExecQuery _
    ("Select * from Win32_QuickFixEngineering")
 		objTextFile.Write("Computer,Description,Hot Fix ID,Installation Date,Installed By ")
		objTextFile.WriteLine
For Each objQuickFix in colQuickFixes
		objTextFile.Write(objQuickFix.CSName) & ","
		objTextFile.Write(objQuickFix.Description) & ","
		objTextFile.Write(objQuickFix.HotFixID) & ","
		objTextFile.Write(objQuickFix.InstallDate) & ","
		objTextFile.Write(objQuickFix.InstalledBy) & ","
		objTextFile.WriteLine
Next
objTextFile.Close
End Sub


'On Error Resume Next