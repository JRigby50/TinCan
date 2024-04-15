On Error Resume Next

Const HKEY_LOCAL_MACHINE = &H80000002

'change this value to the IP address or hostname of the machine you need to audit
Dim strComputer

strComputer = InputBox("Inpute The name Computer ")

CALL GenerateReport(strComputer)

'***************
'GenerateReport
SUB GenerateReport(strComputer)

'Script to change a filename using timestamps
strPath = D:\IT\" 'Change the path to appropriate value For output Directory , you can give a share name line \\myserver\myshare
strMonth = DatePart("m", Now())
strDay = DatePart("d",Now())

If Len(strMonth)=1 Then
strMonth = "0" & strMonth
Else
strMonth = strMonth
End If


If Len(strDay)=1 Then
strDay = "0" & strDay
Else
strDay = strDay
End If


strFileName = DatePart("yyyy",Now()) & strMonth & strDay
strFileName = Replace(strFileName,":","")
'***************

'Variable Declarations
Const ForAppending = 8

'***************
'Main Body
On Error Resume Next


'CompName

strComputer =strComputer

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")


'***************
'For INTERNET EXPLORER
Dim strIE
Set objWMIService2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2\Applications\MicrosoftIE")
Set colIESettings = objWMIService2.ExecQuery("Select * from MicrosoftIE_Summary")
For Each strIESetting in colIESettings
strIE= " INTERNET EXPLORER: " & strIESetting.Name & " v" & strIESetting.Version & VBCRLF
Next


'Get Operation System & Processor Information
Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor",,48)
For Each objItem in colItems
CompName = objItem.SystemName
Next

Set objFSO = CreateObject("Scripting.FileSystemObject")
if objFSO.FileExists(strPath & strComputer & "_Audit.txt") then
WScript.Quit
end if

'Set the file location to collect the data
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(strPath & CompName & "_Audit.txt", ForAppending, True)


'***************
'Print HEADER
objTextFile.Write "================================================================" & VBCRLF & VBCRLF
objTextFile.Write " SERVER RESOURCE AUDIT REPORT " & VBCRLF
objTextFile.Write " DATE: " & FormatDateTime(Now(),1) & " " & VBCRLF
objTextFile.Write " TIME: " & FormatDateTime(Now(),3) & " " & VBCRLF & VBCRLF
objTextFile.Write "================================================================" & VBCRLF & VBCRLF & VBCRLF & VBCRLF & VBCRLF


objTextFile.Write "COMPUTER" & VBCRLF
'***************
'Get OPERATING SYSTEM & Processor Information
objTextFile.Write " COMPUTER NAME: " & CompName & VBCRLF

Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor",,48)
For Each objItem in colItems
objTextFile.Write " PROCESSOR: " & objItem.Name & VBCRLF
Next

Set colProcs = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

For Each objItem in colProcs
objTextFile.Write " NUMBER OF PROCESSORS: " & objItem.NumberOfProcessors & VBCRLF & VBCRLF
Next
'---------------------------------------------------------------------------------------------------------------------------------------
'For Logged on user
Set colComputer = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
 
For Each objComputer in colComputer
    objTextFile.Write "Logged-on user: " & objComputer.UserName & VBCRLF
Next

'***************
'Get DOMAIN NAME information
Set colItems = objWMIService.ExecQuery("Select * from Win32_NTDomain")

For Each objItem in colItems
objTextFile.Write " DOMAIN NAME: " & objItem.DomainName & vbCrLf & VBCRLF
Next

'***************
'Get OS Information
Set colSettings = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objOperatingSystem in colSettings
objTextFile.Write " OPERATING SYSTEM: " & objOperatingSystem.Name & VBCRLF
objTextFile.Write " VERSION: " & objOperatingSystem.Version & VBCRLF
objTextFile.Write " SERVICE PACK: " & objOperatingSystem.ServicePackMajorVersion & "." & objOperatingSystem.ServicePackMinorVersion & VBCRLF
Next
objTextFile.Write strIE & VBCRLF & VBCRLF & VBCRLF & VBCRLF





objTextFile.Write "MOTHERBOARD" & VBCRLF 

'***************
'Get Main Board Information
Set colItems = objWMIService.ExecQuery("Select * from Win32_BaseBoard",,48)
For Each objItem in colItems
objTextFile.Write " MAINBOARD MANUFACTURER: " & objItem.Manufacturer & VBCRLF
objTextFile.Write " MAINBOARD PRODUCT: " & objItem.Product & VBCRLF
Next



'***************
'Get BIOS Information
Set colItems = objWMIService.ExecQuery("Select * from Win32_BIOS",,48)
For Each objItem in colItems
objTextFile.Write " BIOS MANUFACTURER: " & objItem.Manufacturer & VBCRLF
objTextFile.Write " BIOS VERSION: " & objItem.Version & VBCRLF & VBCRLF & VBCRLF & VBCRLF & VBCRLF
Next

objTextFile.Write "MEMORY" & VBCRLF 

'***************
'Get Total Physical memory
Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objComputer in colSettings
objTextFile.Write " TOTAL PHYSICAL RAM: " & Round((objComputer.TotalPhysicalMemory/1000000000),4) & " GB" & VBCRLF
Next

objTextFile.Write " " & VBCRLF & VBCRLF & VBCRLF & VBCRLF & "PARTITIONS" & VBCRLF 

'***************
'Get Logical Disk Size and Partition Information
Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType = 3")
For Each objDisk in colDisks
intFreeSpace = objDisk.FreeSpace
intTotalSpace = objDisk.Size
pctFreeSpace = intFreeSpace / intTotalSpace
objTextFile.Write " DISK " & objDisk.DeviceID & " (" & objDisk.FileSystem & ") " & Round((objDisk.Size/1000000000),4) & " GB ("& Round((intFreeSpace/1000000000)*1.024,4) & " GB Free Space)" & VBCRLF
Next

objTextFile.Write " " & VBCRLF & VBCRLF & VBCRLF & VBCRLF & "NETWORK" & VBCRLF 

'***************
'Get NETWORK ADAPTERS information
Dim strIP, strSubnet, strDescription

Set colNicConfigs = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

For Each objNicConfig In colNicConfigs
'Assign description values to variable
strDescription=objNicConfig.Description

For Each strIPAddress In objNicConfig.IPAddress
'Assign IP Address to variable
strIP=strIPAddress 

For Each strIPSubnet In objNicConfig.IPSubnet
'Assign Subnet to variable
strSubnet = strIPSubnet
Next

objTextFile.Write " NETWORK ADAPTER: " & strDescription & VBCRLF
objTextFile.Write " IP ADDRESS: " & strIP & VBCRLF
objTextFile.Write " SUBNET MASK: " & strSubnet & VBCRLF & VBCRLF

Next

Next

Set colNicConfigs =NOTHING

'***************

objTextFile.Write " " & VBCRLF & VBCRLF & VBCRLF & VBCRLF & "APPLICATION" & VBCRLF 

Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

For Each subkey In arrSubKeys
strSubKeyPath = strKeyPath & "\" & subkey

strString = "DisplayName"
objReg.GetStringValue HKEY_LOCAL_MACHINE, strSubKeyPath, strString, strDisplayName

strString = "DisplayVersion"
objReg.GetStringValue HKEY_LOCAL_MACHINE, strSubKeyPath, strString, strDisplayVersion

strDisplayName=Trim(strDisplayName)
strDisplayVersion=Trim(strDisplayVersion)
If strDisplayName <> "" And strDisplayVersion <> "" Then
objTextFile.Write " " & strDisplayName & " " & strDisplayVersion & VBCRLF
End If
Next

'***************

'Close text file after writing logs

objTextFile.Write VbCrLf
objTextFile.Close

'Clean Up

SET colIESettings=NOTHING
SET colItems=NOTHING
SET colSettings=NOTHING
SET colDisks=NOTHING
SET AdapterSet=NOTHING
SET objWMIService=NOTHING
SET objWMIService2=NOTHING
SET objFSO=NOTHING
SET objTextFile=NOTHING

END SUB
'***************
Function HostOnline(strComputername)

Set strTempFolder = objFso.GetSpecialFolder(TEMPFOLDER)
strTempFile = objFso.GetTempName
strTempFile = strTempFolder & "\" & strTempFile

objShell.Run "cmd /c ping -n 2 -l 8 " & strComputername & ">" & strTempFile,0,True

Set objFile = objFso.GetFile(strTempFile)
set objTS = objFile.OpenAsTextStream(ForReading)
do While objTS.AtEndOfStream <> True
strReturn =objTS.ReadLine
If instr(strReturn, "Reply")>0 then
HostOnline = True
Exit Do
End If
Loop

objTS.Close
objFile.delete
End Function