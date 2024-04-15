'Server Data Collector

Set objNetwork = CreateObject("Wscript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strDirectory1 = "c:\AdminReports"
strDirectory2 = "c:\AdminReports\html"
strComputer = objNetwork.ComputerName
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strcomputer & "\root\cimv2")
'*****Create Folders*************************************************************************************
If objFSO.FolderExists(strDirectory1) Then
   Set objFolder = objFSO.GetFolder(strDirectory1)
Else
   Set objFolder = objFSO.CreateFolder(strDirectory1)
End If
If objFSO.FolderExists(strDirectory2) Then
   Set objFolder = objFSO.GetFolder(strDirectory2)
Else
   Set objFolder = objFSO.CreateFolder(strDirectory2)
End If
'*****BIOS*******************************************************************************************
Set objTextFile = objFSO.CreateTextFile("c:\AdminReports\" & strcomputer & " BIOS.csv", True)
On Error Resume Next
		objTextFile.WriteLine "BIOS" '& vbcrlf
Set colBIOS = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_BIOS")
    objTextFile.WriteLine"Computer name,Manufacturer,Serial number,BIOS Version"
For Each objBIOS in colBIOS
	     objTextFile.Write strcomputer & ","
         objTextFile.Write objBIOS.Manufacturer & ","
         objTextFile.Write objBIOS.SerialNumber & ","
         objTextFile.WriteLine objBIOS.SMBIOSBIOSVersion
Next
objTextFile.Close
'*****Drives***********************************************************************************************
  'Analyze the local Drives and determine if they need to be De-Fragmented
Set objTextFile = objFSO.CreateTextFile("c:\AdminReports\" & strComputer & " Drive Analysis.csv", True)
On Error Resume Next
		objTextFile.WriteLine "Defrag Analysis"
Set colVolumes = objWMIService.ExecQuery("Select * from Win32_Volume")
objTextFile.WriteLine "Volume Name,Volume Size,Used Space,Free Space,Free Space %,Page File Size,Total Folders,Total Files,Average File Size,Average Fragments/File,Total Fragmented Files,Total Page File Fragments,Total % Fragmentation,Needs Defrag? "
For Each objVolume in colVolumes
    errResult = objVolume.DefragAnalysis(blnRecommended, objReport)
    If errResult = 0 Then
        objTextFile.Write (objReport.VolumeName) & ","
        objTextFile.Write (objReport.VolumeSize) & ","
    	objTextFile.Write (objReport.UsedSpace) & ","
    	objTextFile.Write (objReport.FreeSpace) & ","
        objTextFile.Write (objReport.FreeSpacePercent) & ","
        objTextFile.Write (objReport.PageFileSize) & ","
    	objTextFile.Write (objReport.TotalFolders) & ","
    	objTextFile.Write (objReport.TotalFiles) & ","
    	objTextFile.Write (objReport.AverageFileSize) & ","
        objTextFile.Write (objReport.AverageFragmentsPerFile) & ","
        objTextFile.Write (objReport.TotalFragmentedFiles) & ","
        objTextFile.Write (objReport.TotalPageFileFragments) & ","
        objTextFile.Write (objReport.TotalPercentFragmentation) & ","
        If blnRecommended = True Then
           objTextFile.Write "Yes"
        Else
           objTextFile.Write "No"
        End If
        objTextFile.WriteLine
    End If
Next
objTextFile.Close
'*****Event Logs*******************************************************************************************
On error resume Next
Set dtmStartDate = CreateObject("WbemScripting.SWbemDateTime")
Set dtmEndDate = CreateObject("WbemScripting.SWbemDateTime")
DateToCheck = Date - 1
dtmEndDate.SetVarDate Date, True
dtmStartDate.SetVarDate DateToCheck, True
'Application Log
Set objTextFile = objFSO.CreateTextFile("c:\AdminReports\" & strcomputer & " AppLog.csv", True)
Set colLoggedEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = 'Application' and TimeWritten >= '" _ 
        & dtmStartDate & "' and TimeWritten < '" & dtmEndDate & "'")
	objTextFile.Write("Event Type,Time Written,Source,Category,Event,User,Computer")
	objTextFile.WriteLine
For Each objEvent in colLoggedEvents
    objTextFile.Write (objEvent.Type) & ","
    objTextFile.Write (objEvent.TimeWritten) & ","
    objTextFile.Write (objEvent.SourceName) & ","
    objTextFile.Write (objEvent.CategoryString) & ","
    objTextFile.Write (objEvent.Eventcode) & ","
    objTextFile.Write (objEvent.User) & ","
    objTextFile.Write (objEvent.ComputerName) & ","
    objTextFile.WriteLine
Next
'System Log
Set objTextFile = objFSO.CreateTextFile("c:\AdminReports\" & strcomputer & " SysLog.csv", True)
Set colLoggedEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = 'System' and TimeWritten >= '" _ 
        & dtmStartDate & "' and TimeWritten < '" & dtmEndDate & "'")
	objTextFile.Write("Event Type,Time Written,Source,Category,Event,User,Computer")
	objTextFile.WriteLine
For Each objEvent in colLoggedEvents
    objTextFile.Write (objEvent.Type) & ","
    objTextFile.Write (objEvent.TimeWritten) & ","
    objTextFile.Write (objEvent.SourceName) & ","
    objTextFile.Write (objEvent.CategoryString) & ","
    objTextFile.Write (objEvent.EventCode) & ","
    objTextFile.Write (objEvent.User) & ","
    objTextFile.Write (objEvent.ComputerName) & ","
    objTextFile.WriteLine
Next	
objTextFile.Close
'Security Log
Set objTextFile = objFSO.CreateTextFile("c:\AdminReports\" & strcomputer & " SecLog.csv", True)
Set colLoggedEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = 'Security' and TimeWritten >= '" _ 
        & dtmStartDate & "' and TimeWritten < '" & dtmEndDate & "'")
	objTextFile.Write("Event Type,Time Written,Source,Category,Event,User,Computer")
	objTextFile.WriteLine
For Each objEvent in colLoggedEvents
    objTextFile.Write (objEvent.Type) & ","
    objTextFile.Write (objEvent.TimeWritten) & ","
    objTextFile.Write (objEvent.SourceName) & ","
    objTextFile.Write (objEvent.CategoryString) & ","
    objTextFile.Write (objEvent.EventCode) & ","
    objTextFile.Write (objEvent.User) & ","
    objTextFile.Write (objEvent.ComputerName) & ","
    objTextFile.WriteLine
Next
objTextFile.Close
'*****OS******************************************************************************************************
Set objTextFile = objFSO.CreateTextFile("c:\AdminReports\" & strcomputer & " OS.csv", True)
On error resume Next
		objTextFile.WriteLine "OS Version"
		objTextFile.WriteLine "Computer Name,OS,Build,Service Pack,Physical Memory,Free Memory,Windows Directory,Last Reboot"
		Set colItems = objWMIService.ExecQuery("SELECT * " & _
		"FROM Win32_OperatingSystem")
		For Each objItem In colItems		
		 objTextFile.Write objItem.csname & ","
		 objTextFile.Write objItem.Caption & ","
		 objTextFile.Write objItem.BuildNumber & ","
		 objTextFile.Write objItem.CSDVersion & ","
		 objTextFile.Write objItem.TotalVisibleMemorySize & ","
		 objTextFile.Write objItem.FreePhysicalMemory & ","  
		 objTextFile.Write objItem.WindowsDirectory & ","
		 objTextFile.Write objItem.LastBootUpTime & ","
		 objTextFile.WriteLine
Next
objTextFile.Close
'*****Patches***************************************************************************************************
On error resume Next
Set objTextFile = objFSO.CreateTextFile("c:\AdminReports\" & strcomputer & " Hot Fixes.csv", True)
Set colQuickFixes = objWMIService.ExecQuery _
    ("Select * from Win32_QuickFixEngineering")
	objTextFile.Write("Computer,Description,Hot Fix ID,Installation Date,Installed By ")
	objTextFile.WriteLine
For Each objQuickFix in colQuickFixes
If objQuickFix.HotFixId = "File 1" Then
SetBS=1
Else
		objTextFile.Write(objQuickFix.CSName) & ","
		objTextFile.Write(objQuickFix.Description) & ","
		objTextFile.Write(objQuickFix.HotFixID) & ","
		objTextFile.Write(objQuickFix.InstallDate) & ","
		objTextFile.Write(objQuickFix.InstalledBy) & ","
		objTextFile.WriteLine
End If
Next
objTextFile.Close	
'*****Printers********************************************************************************************
On error resume Next
Set objTextFile = objFSO.CreateTextFile("C:\AdminReports\" & strcomputer & " Printers.csv", True)
Set colInstalledPrinters = objWMIService.ExecQuery ("Select * from Win32_Printer")
objTextFile.Write("Name,Port,Driver,Comments,Share Name")
		objTextFile.WriteLine
For Each objPrinter in colInstalledPrinters
		objTextFile.Write(objPrinter.name) & ","
		objTextFile.Write(objPrinter.portname) & ","
		objTextFile.Write(objPrinter.drivername) & ","
		objTextFile.Write(objPrinter.comment) & ","
		objTextFile.Write(objPrinter.sharename) & ","
		objTextFile.WriteLine
Next
objTextFile.Close	
'*****Services*********************************************************************************************
On error resume Next
Set objTextFile = objFSO.CreateTextFile("c:\AdminReports\" & strcomputer & " Services.csv", True)
'Services set to run Automaticaly
		objTextFile.WriteLine
		objTextFile.WriteLine "Automatic Services, , , ,"
		objTextFile.WriteLine
		objTextFile.WriteLine("Display Name,Service Name,Service State,Start Mode,Process ID ")
		objTextFile.WriteLine
Set colAutomaticServices = objWMIService.ExecQuery _
    ("Select * from Win32_Service Where startMode = 'Auto'")
For Each objService in colAutomaticServices
		objTextFile.Write(objService.DisplayName) & ","
		objTextFile.Write(objService.Name) & ","
		objTextFile.Write(objService.State) & ","
		objTextFile.Write(objService.StartMode) & ","
		objTextFile.Write(objService.ProcessID) & ","
		objTextFile.WriteLine   
Next
'Services set to run Manually
		objTextFile.WriteLine
		objTextFile.WriteLine "Manual Services, , , ,"
		objTextFile.WriteLine
		objTextFile.WriteLine("Display Name,Service Name,Service State,Start Mode,Process ID ")
		objTextFile.WriteLine
Set colManualServices = objWMIService.ExecQuery _
    ("Select * from Win32_Service Where startMode = 'Manual'")
For Each objService in colManualServices
		objTextFile.Write(objService.DisplayName) & ","
		objTextFile.Write(objService.Name) & ","
		objTextFile.Write(objService.State) & ","
		objTextFile.Write(objService.StartMode) & ","
		objTextFile.Write(objService.ProcessID) & ","
		objTextFile.WriteLine 
Next
'Disabled Services
		objTextFile.WriteLine
		objTextFile.WriteLine "Disabled Services, , , ,"
		objTextFile.WriteLine
		objTextFile.WriteLine("Display Name,Service Name,Service State,Start Mode,Process ID ")
		objTextFile.WriteLine
Set colDisabledServices = objWMIService.ExecQuery _
    ("Select * from Win32_Service Where startMode = 'Disabled'")
For Each objService in colDisabledServices
		objTextFile.Write(objService.DisplayName) & ","
		objTextFile.Write(objService.Name) & ","
		objTextFile.Write(objService.State) & ","
		objTextFile.Write(objService.StartMode) & ","
		objTextFile.Write(objService.ProcessID) & ","
		objTextFile.WriteLine
Next
objTextFile.Close
'*****Shares****************************************************************************************************
On error resume Next
Set objTextFile = objFSO.CreateTextFile("c:\AdminReports\" & strcomputer & " Shares.csv", True)
Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")
objTextFile.Write("Allow Maximum,Caption,Maximum Allowed,Name,Path")
objTextFile.WriteLine
For each objShare in colShares
If objShare.Type<>1 Then
objTextFile.Write(objShare.AllowMaximum) & ","
objTextFile.Write(objShare.Caption) & ","
objTextFile.Write(objShare.MaximumAllowed) & ","
objTextFile.Write(objShare.Name) & ","
objTextFile.Write(objShare.Path) & ","
objTextFile.WriteLine
End If
Next
objTextFile.Close
'*****Software*********************************************************************************************
On error resume next
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
Set objTextFile = objFSO.CreateTextFile("c:\AdminReports\" & strcomputer & " Installed Software.csv", True)
objTextFile.WriteLine "Installed Software"
objTextFile.WriteLine
strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry1a = "DisplayName"
strEntry1b = "QuietDisplayName"
Set objReg = GetObject("winmgmts://" & strComputer & _
 "/root/default:StdRegProv")
objReg.EnumKey HKLM, strKey, arrSubkeys
For Each strSubkey In arrSubkeys
  intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, _
   strEntry1a, strValue1)
  If intRet1 <> 0 Then
    objReg.GetStringValue HKLM, strKey & strSubkey, _
     strEntry1b, strValue1
  End If
  If strValue1 <> "" Then
objTextFile.WriteLine strValue1
  End If
Next
strKey2 = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry2a = "DisplayName"
strEntry2b = "QuietDisplayName"
Set objReg = GetObject("winmgmts://" & strComputer & _
 "/root/default:StdRegProv")
objReg.EnumKey HKLM, strKey2, arrSubkeys
For Each strSubkey In arrSubkeys
  intRet2 = objReg.GetStringValue(HKLM, strKey2 & strSubkey, _
   strEntry2a, strValue2)
  If intRet2 <> 0 Then
    objReg.GetStringValue HKLM, strKey & strSubkey, _
     strEntry2b, strValue2
  End If
  If strValue2 <> "" Then
objTextFile.WriteLine strValue2
  End If
Next
objTextFile.Close
