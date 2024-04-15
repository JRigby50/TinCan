On Error Resume Next

Set objShell = WScript.CreateObject ("WScript.shell")
Set objFSO = WScript.CreateObject ("Scripting.FileSystemObject")
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Set objComputerName = CreateObject("Wscript.Network") 
strComputer = objComputerName.ComputerName 

strFilePath = "C:\Temp\"
strDefragLog = strFilePath & Replace((strComputer & "_Defrag_" & Date() & ".log"),"/","-")

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")
	For Each objDisk in colDisks	
		If objDisk.DriveType = 3 Then

			intDriveCount = intDriveCount + 1
		
			strDriveLog = """" & strFilePath & "TempLog_" & Mid(objDisk.DeviceID,1,1) & """"

			If intDriveCount = 1 Then
				strBeginTimes = Now()
			Else
				strBeginTimes = strBeginTimes & "," & Now()
			End If

			strDriveLetter = objDisk.DeviceID
			
			strcmd = "defrag.exe -v " & strDriveLetter & " >" & strDriveLog
			intReturn = objShell.Run ("cmd /C" & strcmd,0,True)
			
			Select Case intReturn
				Case 0
					strReturnMessage = "Success"
				Case 1
					strReturnMessage = "Access Denied"
				Case 2
					strReturnMessage = "Not Supported"
				Case 3
					strReturnMessage = "Volume dirty bit is set"
				Case 4
					strReturnMessage = "Not enough free space"
				Case 5
					strReturnMessage = "Corrupt Master File Table detected"
				Case 6
					strReturnMessage = "Call cancelled"
				Case 7
					strReturnMessage = "Call cancellation request too late"
				Case 8
					strReturnMessage = "Defrag engine is already running"
				Case 9
					strReturnMessage = "Unable to connect to defrag engine"
				Case 10
					strReturnMessage = "Defrag engine error"
				Case 11
					strReturnMessage = "Unknown Error"
			End Select
			
			
			If intReturn <> 0 Then
				intDriveFailCount = intDriveFailCount + 1
				If intDriveFailCount = 1 Then
					strDriveFail = objDisk.DeviceID
					strFailMessage = strReturnMessage
				Else
					strDriveFail = strDriveFail & objDisk.DeviceID
					strFailMessage = strFailMessage & "," & strReturnMessage
				End If
			End If

			If intDriveCount = 1 Then
    			strDrivesList = objDisk.DeviceID
				strEndTimes = Now()
				strTempLogs = Replace(strDriveLog, """","")
			Else
    			strDrivesList = strDrivesList & "," & objDisk.DeviceID
				strEndTimes = strEndTimes & "," & Now()
				strTempLogs = strTempLogs & "," & Replace(strDriveLog, """","")
    		End If
		End If
	Next	'objDisk

If Len(strDriveFail) > 0 Then
	If Len(strDriveFail) > 1 Then
		arrFailMessage = Split(strFailMessage, ",")
	Else
		arrFailMessage = strFailMessage
	End If
End If


If intDriveCount > 1 Then
	arrDrivesList = Split(strDrivesList , ",")
	arrEndTimes = Split(strEndTimes , ",")
	arrBeginTimes = Split(strBeginTimes , ",")
	arrTempLogs = Split(strTempLogs, ",")
End If


Set objLogFile = objFSO.OpenTextFile (strDefragLog, ForAppending, True)
objLogFile.WriteLine strComputer & " " & Date() & vbcrlf

If intDriveCount > 1 Then
	For i = 0 To Ubound(arrDrivesList)
		Set objTempFile = objFSO.OpenTextFile (arrTempLogs(i), ForReading, True)
		strTempFileAll = objTempFile.ReadAll
		objLogFile.WriteLine "---Drive " & arrDrivesList(i) & " Log Section---" & vbcrlf
		objLogFile.Write strTempFileAll
		objTempFile.Close
		If InStr(strDriveFail,arrDrivesList(i)) <> 0 Then
			objLogFile.WriteLine "Defragmentation FAILURE - Drive " & arrDrivesList(i) & " - " & arrFailMessage(intFailCounter) & vbcrlf
			intFailCounter = intFailCounter + 1
		Else	
			objLogFile.WriteLine vbcrlf & "Defragmentation of " &  arrDrivesList(i) & " completed successfully:"
			objLogFile.WriteLine "Start time" & vbTab & "= " & arrBeginTimes(i)
			objLogFile.WriteLine "End time" & vbTab & "= " & arrEndTimes(i) & vbcrlf & vbcrlf
		End If
		objFSO.DeleteFile (arrTempLogs(i))
	Next 'i
Else
	If InStr(strDriveFail, strDrivesList) <> 0 Then
		objLogFile.WriteLine vbcrlf & "Defragmentation FAILURE - Drive " & strDrivesList & " - " & strFailMessage & vbcrlf
	Else
		objLogFile.WriteLine vbcrlf & "Defragmentation of " &  strDrivesList & " completed successfully:"
		objLogFile.WriteLine "Start time" & vbTab & "= " & strBeginTimes
		objLogFile.WriteLine "End time" & vbTab & "= " & strEndTimes
	End If
End If

