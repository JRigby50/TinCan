Option Explicit
Dim objFSO, objTextFile, objReg, objServerList
Dim strServer, strKeyPath, strFont
Dim arrFonts
Dim bolPresent
Const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const OverwriteExisting = True
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile("E:\IT\CalibriInstalled.txt", ForWriting)
objTextFile.WriteLine "Calibri is Installed on:"
objTextFile.WriteLine
Set objServerList = objFSO.OpenTextFile("E:\IT\PrintServers.txt", ForReading)

Do Until objServerList.AtEndOfStream
	strServer = objServerList.ReadLine
	bolPresent = False
	On Error Resume Next
	'strServer = "."
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strServer & "\root\default:StdRegProv")
	strKeyPath = "Software\Microsoft\Windows NT\CurrentVersion\Fonts"
	objReg.EnumValues HKEY_LOCAL_MACHINE, strKeyPath,arrFonts
	For Each strFont in arrFonts
		If InStr (strFont,"Calibri") Then
		bolPresent = True	
		End If
	Next
	If bolPresent Then
		objTextFile.WriteLine strServer
	End If
Loop

objServerList.Close
objTextFile.Close