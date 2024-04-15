'~~[author]~~
'Alex Woolsey
'~~[/author]~~

'~~[emailAddress]~~
'Alex_Woolsey@hotmail.com
'~~[/emailAddress]~~

'~~[scriptType]~~
'vbscript
'~~[/scriptType]~~

'~~[subType]~~
'Misc
'~~[/subType]~~

'~~[keywords]~~
'A Backup & Clear Event Log Script that reads Excel for the Remote Machine Name / Username & Password.
'~~[/keywords]~~

'~~[comment]~~
'To use the script you need to create an Excel Spreadsheet with 1 of the workbooks labelled as "Servers". Then in Column A type in the name or IP of your target machine. In Column B type in the username (either machine\user or domain\user) then in column C type in the password for that account. If your own account has full domain admin permissions you only need to type in the name of the target machine in Column A - leave B & C empty.

The script will then go off & backup & clear the event log on the remote machine, it will tell you along the way which machine it is processing & report on folders created etc etc (So please use CSCRIPT!) It will also log an event on the local & remote machine advising where the backup was placed & it will then copy it back to your own machine.
If folders don't exist on local or remote machines they will be created - you can change the path in the script if you like. At present it creates a C:\EventLogs\todaysdate\machinename folder & puts the logs in there..  It also creates a C:\EventLogs on the target machine & puts the logs in there...   I've remmed out the deletion line but feel free to play around. I have objexcel.quit in twice because it doesn't seem to close off properly on my own machine!

Many thanks go to the other scripters about who already had backup & clear event log scripts, I did fudge around with some that others had supplied.
If anyone makes it even better then please post it back !
Kind Regards
Alex
'~~[/comment]~~

'~~[script]~~
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\
'\\	Backup & Clear Event Logs & Move to Local Machine
'\\	Alex Woolsey - 04/07/05
'\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Set objNetwork = CreateObject("WScript.Network")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
Set objExcel = CreateObject("Excel.Application")

myCur = WshShell.CurrentDirectory
myFol = "C:\EventLogs\"
inputfile = myCur & "\" & "Servers.xls"
strBackupName = left(now(),2) & "_" & mid(now(),4,2) & "_" & mid(now(),7,4)
NewFol = myFol & strBackupName
Set objWorkbook = objExcel.Workbooks.Open(inputfile)

If not (FSO.FolderExists(myFol)) Then
	FSO.CreateFolder(myFol)
Wscript.Echo VbTab & myFol & " Created.."
End If
If not (FSO.FolderExists(NewFol)) Then
	FSO.CreateFolder(NewFol)
Wscript.Echo VbTab & NewFol & " Created.."
End If

i = 1
Do until objWorkbook.Sheets("Servers").Cells(i,1).Value = ""
strComputer = objWorkbook.Sheets("Servers").Cells(i,1).Value
strUser = objWorkbook.Sheets("Servers").Cells(i,2).Value
strPwd = objWorkbook.Sheets("Servers").Cells(i,3).Value

If strUser = "" Then
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!//" & strComputer & "/root/cimv2")
	Wscript.Echo VbCrLf & "Processing... " & strComputer & vbTab & "Using Windows Impersonate Credentials..." & VbCrLf
		If Err <> 0 Then
			Wscript.Echo strComputer & VbTab & Err.Description & VbCrLf
		Err.Clear
		i = i + 1
		End If
Else
	Set objWMILocator = CreateObject("WbemScripting.SWbemLocator")
	Set objWMIService = objWMILocator.ConnectServer(strComputer,"root\cimv2",strUser,strPwd)
	Wscript.Echo VbCrLf & "Processing... " & strComputer & vbTab & "Using : " & strUser & VbCrLf
		If Err <> 0 Then
			Wscript.Echo strComputer & VbTab & Err.Description & VbCrLf
		Err.Clear
		i = i + 1
		End If
End If
zFol = NewFol & "\" & strComputer
If not (fso.FolderExists(zFol)) Then
	fso.CreateFolder(zFol)
Wscript.Echo VbTab & zFol & " Created.."

End If

sDestination = "\\" & strComputer & "\c$\" & "EventLogs"
If Not fso.FolderExists(sDestination) Then
	Dim snewfolder1
	snewfolder1 = fso.CreateFolder(sDestination)
Wscript.Echo VbTab & snewfolder1 & " Created.." & VbCrLf
End If

oApplog = strComputer & "_" & strBackupName & "_App.evt"
oSyslog = strComputer & "_" & strBackupName & "_Sys.evt"
oSeclog = strComputer & "_" & strBackupName & "_Sec.evt"
'*************************************************************************************** Application Log

Set colLogFiles = objWMIService.ExecQuery _
    ("Select * from Win32_NTEventLogFile where LogFileName='Application'")
For Each objLogfile in colLogFiles
    objLogFile.BackupEventLog(myFol & oApplog)
    objLogFile.ClearEventLog()
    WshShell.LogEvent 0, "The Application log was backed up to " & myFol & oApplog, strComputer
    WshShell.LogEvent 0, "The Application log from " & strComputer & " was backed up to " & myFol &  oApplog
	Wscript.Echo VbTab & oApplog & " Done.."
Next

'*************************************************************************************** System Log

Set colLogFiles = objWMIService.ExecQuery _
    ("Select * from Win32_NTEventLogFile where LogFileName='System'")
For Each objLogfile in colLogFiles
    objLogFile.BackupEventLog(myFol & oSyslog)
    objLogFile.ClearEventLog()
    WshShell.LogEvent 0, "The System log was backed up to " & myFol & oSyslog, strComputer
    WshShell.LogEvent 0, "The System log from " & strComputer & " was backed up to " & myFol & oSyslog
	Wscript.Echo VbTab & oSyslog & " Done.."
Next

'*************************************************************************************** Security Log

Set colLogFiles = objWMIService.ExecQuery _
    ("Select * from Win32_NTEventLogFile where LogFileName='Security'")
For Each objLogfile in colLogFiles
    objLogFile.BackupEventLog(myFol & oSeclog)
    objLogFile.ClearEventLog()
    WshShell.LogEvent 0, "The Security log was backed up to " & myFol & oSeclog, strComputer
    WshShell.LogEvent 0, "The Security log from " & strComputer & " was backed up to " & myFol & oSeclog
	Wscript.Echo VbTab & oSeclog & " Done.."
Next
wscript.sleep 5000

'Set the path to the remote logs
LastPath = sDestination & "\"

If FSO.FolderExists (LastPath) Then
        FSO.CopyFile LastPath & "*.evt" , zFol , True
        wscript.echo VbCrLf & strComputer & " Logs Moved To: " & zFol
        'FSO.DeleteFolder sDestination, True
	Wscript.Echo strComputer & " " & sDestination & " Deleted.." & VbCrLf
	Wscript.Echo "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
      End If
i = i + 1
Loop
Wscript.Echo "All Finished...."
objExcel.ActiveWorkbook.Close True
objExcel.quit
objExcel.quit
Set objExcel = Nothing
'~~[/script]~~
