'Script to Copy Data Files to a Central Repository.
'**************************************************************************
'Change Drive Leter For Deployment
'**************************************************************************
Option Explicit
'Define Constants And Variables
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const OverwriteExisting = TRUE
Dim objFSO, objServerList, objLogFile, objDrive
Dim strServerName, bolSuccess, strNow
strNow = Now
'Create Basic Objects
Set objFSO = WScript.CreateObject("Scripting.Filesystemobject")
Set objLogFile = objFSO.OpenTextFile("E:\dcs\DHCP_Collecter_Log.txt", ForAppending)
objLogFile.WriteLine
objLogFile.WriteLine strNow
'**************************************************************************************************
'Cycle through the Server Names in DHCPServers.txt and Backup the Backups
Set objServerList = objFSO.OpenTextFile("E:\dcs\DHCPServers.txt", ForReading)
Do Until objServerList.AtEndOfStream
strServerName = objServerList.ReadLine
 On Error Resume Next
Set objDrive = WScript.CreateObject("WScript.Network")
objDrive.MapNetworkDrive "Z:", "\\" & strServerName & "\DCS"
'Copy DHCP Export File
'Clean up old file if it exists
If objFSO.FileExists("E:\dcs\Export\" & strServerName & "_Export.txt")Then
objFSO.DeleteFile("E:\dcs\Export\" & strServerName & "_Export.txt")
End If
bolSuccess = False
objFSO.CopyFile "Z:\export.txt", "E:\dcs\Export\" & strServerName & "_Export.txt", OverwriteExisting
' Check for Success or Error and make the appropriate entry in the DHCP_Collecter_Log.txt
'MsgBox Err
If Err = 0 Then
bolSuccess = True
End If
'Log Task Status
If bolSuccess Then
objLogFile.WriteLine "DHCP Export File Copied for server " & strServerName
Else
objLogFile.WriteLine "DHCP Export File Not Copied for server " & strServerName
End If
'**************************************************************************************************
'Copy DHCP Dump File
'Clean up old file if it exists
If objFSO.FileExists("E:\dcs\Dump\" & strServerName & "_Dump.txt") Then
objFSO.DeleteFile("E:\dcs\Dump\" & strServerName & "_Dump.txt")
End If
bolSuccess = False
objFSO.CopyFile "Z:\dump.txt", "E:\dcs\Dump\" & strServerName & "_Dump.txt", OverwriteExisting
' Check for Success or Error and make the appropriate entry in the DHCP_Collecter_Log.txt
If Err = 0 Then
bolSuccess = True
End If
'MsgBox Err
'Log Task Status
If bolSuccess Then
objLogFile.WriteLine "DHCP Dump File Copied for server " & strServerName
Else
objLogFile.WriteLine "DHCP Dump File Not Copied for server " & strServerName
End If
objDrive.RemoveNetworkDrive "Z:"
Loop
'**************************************************************************************************
'All done now cleanup.
objServerList.Close
objLogFile.Close
'******************