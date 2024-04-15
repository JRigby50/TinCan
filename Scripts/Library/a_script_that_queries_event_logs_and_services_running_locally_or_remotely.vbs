'~~[author]~~
'Tunji M Taiwo
'~~[/author]~~

'~~[emailAddress]~~
'tujayboi@verizon.net
'~~[/emailAddress]~~

'~~[scriptType]~~
'vbscript
'~~[/scriptType]~~

'~~[subType]~~
'Misc
'~~[/subType]~~

'~~[keywords]~~
'Event Logs, Services, HTA
'~~[/keywords]~~

'~~[comment]~~
'A script that queries Event Logs and Services running Locally or Remotely.
'~~[/comment]~~

'~~[script]~~
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="ProgId" content="FrontPage.Editor.Document">
<META HTTP-EQUIV="Expires" CONTENT="0">
<title>EVENT VIEWER</title>
<HTA:APPLICATION 
     NAVIGABLE = no
     ID="DRecoveryISP"
     APPLICATIONNAME="Disaster Recovery"
     SCROLL="no"
     BORDERSTYLE="normal"
     SINGLEINSTANCE="yes"
     WINDOWSTATE="maximize"
     BORDER="thin"
     CAPTION="yes"
     MAXIMIZEBUTTON="no"
	 MINIMIZEBUTTON="no"
	 SHOWINTASKBAR="no"
	 SYSMENU="yes"
>
</head>
'--------------------------------------
'Desgined by Tunji M Taiwo 03/28/2006
'--------------------------------------

<SCRIPT LANGUAGE="VBScript">
Sub backup
'*************************************************************************
' Declare Variables
'*************************************************************************

Dim dtmThisDay,dtmThisMonth,dtmThisYear
Dim strBackUpDate,strComputerName,strLogName,strLine,errBackupLog
Dim arrEvent,colLogFile,objExplorer,objDocument
Dim objFSO,objFile,objWMIService,objLogFile,objShell,pathName,extension

Const TIMEOUT = 3
Const ForReading = 1

'*************************************************************************
' Set Value for Variables
'*************************************************************************

dtmThisDay = Day(Now)
dtmThisMonth = Month(Now)
dtmThisYear = Year(Now)

strLogName = ("Security")

'*************************************************************************
'Sets Backup date portion of file name
'*************************************************************************
strComputerName = StoreEntry.value
strBackupDate = InputBox("Enter Date in MM-DD-YYYY Format") 'Date
pathName =  destination.value 'InputBox("Enter destination for backup","Destination","C:\Eventlogs\")
extension = ext.value 'InputBox("Enter extension","Extension",".evt") 

   
'*************************************************************************
' Backup and Clear SYSTEM Event Log
'*************************************************************************
If systemCheck.Checked Then
display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>Processing......</font>"

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate,(Backup, Security)}!\\" & _
strComputerName & "\root\cimv2")

Set colLogFile = objWMIService.ExecQuery _
("Select * from Win32_NTEventLogFile where LogFileName ='System'")

    For Each objLogfile in colLogFile
    errBackupLog = objLogFile.BackupEventLog _
   (pathName & strBackupDate & "-" & strComputerName & "-"_
    & objLogFile.LogFileName & extension)
        If errBackupLog <> 0 Then 
        
        'dim shell
        'Set objShell = WScript.CreateObject ("Wscript.Shell")
            'objShell.Popup "The System event log could not be backed up."_
            '& vbCrLf & vbCrLf & "File may already exsit"_
            '& vbCrLf & vbCrLf & "Insure Path is correct", TIMEOUT
            
                        Call HTASleep(5)

            
            Set objShell = Nothing
    
        Else
		'use to clear log if needed
        'objLogFile.ClearEventLog()

        End If
    
Next
'display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>Completed</font>"

End if

Call HTASleep(10)


'*************************************************************************
' Backup and Clear APPLICATION Event Log
'*************************************************************************
if applicationCheck.Checked Then
display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>Processing......</font>"

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate,(Backup, Security)}!\\" & _
strComputerName & "\root\cimv2")

Set colLogFile = objWMIService.ExecQuery _
("Select * from Win32_NTEventLogFile where LogFileName ='Application'")

    For Each objLogfile in colLogFile
    errBackupLog = objLogFile.BackupEventLog _
    (pathName & strBackupDate & "-" & strComputerName & "-"_
    & objLogFile.LogFileName & extension)

        If errBackupLog <> 0 Then

        
        'dim shell
        'Set objShell = WScript.CreateObject ("Wscript.Shell")
            'objShell.Popup "The System event log could not be backed up."_
            '& vbCrLf & vbCrLf & "File may already exsit"_
            '& vbCrLf & vbCrLf & "Insure Path is correct", TIMEOUT
            
                        Call HTASleep(5)

                            
        Else
		'use to clear log if needed
        'objLogFile.ClearEventLog()

        End If
    
Next
'display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>Completed</font>"

End if

Call HTASleep(10)

'*************************************************************************
' Backup and Clear SECURITY Event Log
'*************************************************************************
If securityCheck.Checked Then
display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>Processing</font>"

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate,(Backup, Security)}!\\" & _
strComputerName & "\root\cimv2")

Set colLogFile = objWMIService.ExecQuery _
("Select * from Win32_NTEventLogFile where LogFileName ='Security'")

    For Each objLogfile in colLogFile
    errBackupLog = objLogFile.BackupEventLog _
(pathName & strBackupDate & "-" & strComputerName & "-"_
    & objLogFile.LogFileName & extension)


            If errBackupLog <> 0 Then

			
        'dim shell
        'Set objShell = WScript.CreateObject ("Wscript.Shell")
            'objShell.Popup "The System event log could not be backed up."_
            '& vbCrLf & vbCrLf & "File may already exsit"_
            '& vbCrLf & vbCrLf & "Insure Path is correct", TIMEOUT
            
                           Call HTASleep(5)

            
            Set objShell = Nothing

    Else
		'use to clear log if needed
        'objLogFile.ClearEventLog()

    End If
    
Next
'display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>Completed</font>"

End if

display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>Completed</font>"

Set objFSO = CreateObject("Scripting.FileSystemObject")

'*************************************************************************
' Clean up Objects
'*************************************************************************

Set objFSO = Nothing
Set objFile = Nothing
Set objExplorer = Nothing
Set objDocument = Nothing
Set objShell = Nothing
Set objLogFile = Nothing
Set objWMIService = Nothing End Sub

'This is what I use to Call a Sleep Procedure'
'-----------------------------------------------------------------------------'
Sub HTASleep(intSeconds)
Dim objShell, strCommand

Set objShell = CreateObject("Wscript.Shell")

strCommand = "%COMSPEC% /c ping -n " & 1 + intSeconds & " 127.0.0.1>nul"
objShell.Run strCommand, 0, True

Set objShell = Nothing
End Sub
'------------------------------------------------------------------------------'

Sub window_onload
Dim shell
Set shell=createobject("wscript.shell")
strMachines = shell.ExpandEnvironmentStrings("%COMPUTERNAME%")
aMachines = split(strMachines, ";")

dateLabel.innerHTML="<font color=red>" & strMachines & "</font>"
self.Focus()
self.ResizeTo 800,600
self.moveto (210)/2, (150)/2
StoreEntry.value=strMachines
'destination.value="C:\Eventlogs\"
'ext.value=".evt"
Bar.document.location.href="Blank.html"
Logtype.innerHTML="Events Logs Controller on"
StoreName.Value=strMachines
ApplicationCheck.Checked=true
'ServiceName.value="%serviceName%"
StartDate.Value=Date
EndDate.value=Date

End Sub

'Checkbox Selections
Sub Pick

strComputer=StoreEntry.Value
StoreName.value=UCase(strComputer)

'-----------------------------------------
'Section for Application Logs
'-----------------------------------------
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    strHTML = "<table border='0' width='100%' id='table1' cellspacing='0'>"
    strHTML = strHTML & "<tr>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #C0C0C0; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>EventType</font></td>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #C0C0C0; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>Source</font></td>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #C0C0C0; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>EventId</font></td>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #C0C0C0; vertical-align:middle' width='0' align='left'><font size='2' face='Arial'>Description</font></td>"
     
 strHTML = strHTML & "</tr>"
      
    If ApplicationCheck.Checked Then
    runbutton.disabled=true
    display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>Processing a Query of the Application Events Log....<p> Please Wait.......</font>"
    Logtype.innerHTML="Querying..."
    Call HTASleep(2)
    Const CONVERT_TO_LOCAL_TIME = True
	Set dtmStartDate = CreateObject("WbemScripting.SWbemDateTime")
	Set dtmEndDate = CreateObject("WbemScripting.SWbemDateTime")
	DateToCheck = CDate(StartDate.value)
	dtmStartDate.SetVarDate DateToCheck, CONVERT_TO_LOCAL_TIME
	dtmEndDate.SetVarDate DateToCheck + 1, CONVERT_TO_LOCAL_TIME
	
	Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where TimeWritten >= '" _ 
        & dtmStartDate & "' and TimeWritten < '" & dtmEndDate & "'")
    
For each objEvent in colEvents
If objEvent.LogFile = "Application" then
	Logtype.innerHTML=" Application Events Logs on a Computer named:"
    strHTML = strHTML & "<tr>"
         If objEvent.Type = "error" Then
			strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: red; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.Type & "</font></td>"
         ElseIf  objEvent.Type = "information" Then
			strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: blue; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.Type & "</font></td>"
         ElseIf  objEvent.Type = "warning" Then
			strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: yellow; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.Type & "</font></td>"
         End If
      	strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.Type & "</font></td>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.SourceName & "</font></td>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.EventCode & "</font></td>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='left'><font size='2' face='Arial'>" & objEvent.Message & "</font></td>"
     strHTML = strHTML & "</tr>"
    
End if   
runbutton.disabled=false
Next
    
Else
        display.document.body.InnerHTML = ""
    End If
    
'-----------------------------------------
'Section for System Logs
'----------------------------------------- 
    If SystemCheck.Checked Then
    runbutton.disabled=true
    display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>Processing a Query of the System Events Log....<p> Please Wait.......</font>"
    Logtype.innerHTML="Querying..."
    Call HTASleep(2)
    Set dtmStartDate = CreateObject("WbemScripting.SWbemDateTime")
	Set dtmEndDate = CreateObject("WbemScripting.SWbemDateTime")
	DateToCheck = CDate(StartDate.value)
	dtmStartDate.SetVarDate DateToCheck, CONVERT_TO_LOCAL_TIME
	dtmEndDate.SetVarDate DateToCheck + 1, CONVERT_TO_LOCAL_TIME
	'strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where TimeWritten >= '" _ 
        & dtmStartDate & "' and TimeWritten < '" & dtmEndDate & "'") 
    
For each objEvent in colEvents
If objEvent.LogFile = "System" then
Logtype.innerHTML=" System Events Logs on a Computer named:"
    strHTML = strHTML & "<tr>"
      If objEvent.Type = "error" Then
			strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: red; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.Type & "</font></td>"
      ElseIf  objEvent.Type = "information" Then
			strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: blue; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.Type & "</font></td>"
      ElseIf  objEvent.Type = "warning" Then
			strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: yellow; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.Type & "</font></td>"
      End If      
      	strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.Type & "</font></td>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.SourceName & "</font></td>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.EventCode & "</font></td>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='left'><font size='2' face='Arial'>" & objEvent.Message & "</font></td>"
     strHTML = strHTML & "</tr>"
End if  
runbutton.disabled=false 
Next
    Else
        display.document.body.InnerHTML = ""

    End If
    
'-----------------------------------------
'Section for Security Logs
'----------------------------------------- 
If SecurityCheck.Checked Then
	Logtype.innerHTML="Querying..."
    runbutton.disabled=true
	display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>Processing a Query of the Security Events Log....<p> Please Wait.......</font>"
	Call HTASleep(2)
    Set dtmStartDate = CreateObject("WbemScripting.SWbemDateTime")
	Set dtmEndDate = CreateObject("WbemScripting.SWbemDateTime")
	DateToCheck = CDate(StartDate.value)
	dtmStartDate.SetVarDate DateToCheck, CONVERT_TO_LOCAL_TIME
	dtmEndDate.SetVarDate DateToCheck + 1, CONVERT_TO_LOCAL_TIME
	Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where TimeWritten >= '" _ 
    & dtmStartDate & "' and TimeWritten < '" & dtmEndDate & "'") 
    
For each objEvent in colEvents
If objEvent.LogFile = "Security" then
	Logtype.innerHTML=" Security Events Logs on a Computer named:"
    strHTML = strHTML & "<tr>"
      If objEvent.Type = "error" Then
			strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: red; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.Type & "</font></td>"
      ElseIf  objEvent.Type = "information" Then
			strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: blue; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.Type & "</font></td>"
      ElseIf  objEvent.Type = "warning" Then
			strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: yellow; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.Type & "</font></td>"
     End If
     
      	strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.Type & "</font></td>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.SourceName & "</font></td>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.EventCode & "</font></td>"
        'strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objEvent.SourceName & "</font></td>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='left'><font size='2' face='Arial'>" & objEvent.Message & "</font></td>"
     strHTML = strHTML & "</tr>"
     strHTML = strHTML & "<p>"  
End if  
runbutton.disabled=false 
Next
    Else
        display.document.body.InnerHTML = ""
    End If

  strHTML = strHTML & "</table>"

    If strHTML <> "" Then
         display.document.body.InnerHTML = strHTML
	End If
runbutton.disabled=false
End Sub  


'-------------------------------------------------------
'List of All Services on computer
'-------------------------------------------------------
Sub listServices
restartC.disabled=true
stopServiceC.disabled=true
startServiceC.disabled=true
listServicesC.disabled=true
strComputer=StoreEntry.Value
StoreName.value=UCase(strComputer)
Logtype.innerHTML=" Querying...."
Display.document.body.innerHTML=""
display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>Processing a Query of all Services....<p> Please Wait.......</font>"
   
Call HTASleep(3)

strHTML = "<table border='0' width='100%' id='table1' cellspacing='0'>"
    strHTML = strHTML & "<tr>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #C0C0C0; vertical-align:left' width='0' align='center'><font size='2' face='Arial'>Service Name</font></td>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #C0C0C0; vertical-align:left' width='0' align='center'><font size='2' face='Arial'>Display Name</font></td>"
        strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #C0C0C0; vertical-align:left' width='0' align='center'><font size='2' face='Arial'>Status</font></td>"
    strHTML = strHTML & "</tr>"

Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\cimv2")
Set colServices = objWMIService.ExecQuery _
    ("Select * From Win32_Service")

For Each objService in colServices
if objService.Name <> ucase(0) then
	Logtype.innerHTML=" Service List on:"
	strHTML = strHTML & "<tr>"
         strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objService.Name & "</font></td>"
         strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & objService.DisplayName & "</font></td>" 'strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #C0C0C0; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>EventId</font></td>"
    	 'if objService.State = Ucase(1) then
    	 'strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: red; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & UCASE(objService.State) & "</font></td>" 'strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #C0C0C0; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>EventId</font></td>"
    	 'elseif objService.State = Ucase(4) then
    	 'strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: blue; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & UCASE(objService.State) & "</font></td>" 'strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #C0C0C0; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>EventId</font></td>"
    	 'end if 
    	 strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #FFFFFF; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>" & UCASE(objService.State) & "</font></td>" 'strHTML = strHTML & "<td style='color: #000000; border-style: outset; border-width: 1px; background-color: #C0C0C0; vertical-align:middle' width='0' align='center'><font size='2' face='Arial'>EventId</font></td>"
    	
    strHTML = strHTML & "</tr>"
    End if   
Next
strHTML = strHTML & "</table>"
display.document.body.InnerHTML=strHtml
restartC.disabled=false
stopServiceC.disabled=false
startServiceC.disabled=false
listServicesC.disabled=false
End Sub


'This Part starts the service
'-------------------------------------------------
Sub startService
restartC.disabled=true
stopServiceC.disabled=true
startServiceC.disabled=true
listServicesC.disabled=true
strComputer=StoreEntry.Value
StoreName.value=UCase(strComputer)
Display.document.body.innerHTML=""
 
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    
Set colServiceList = objWMIService.ExecQuery _
    ("Select * from Win32_Service")

For Each objService in colServiceList
if objService.Name = serviceName.value then
    Logtype.innerHTML="Processing...."
	display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>" & "Starting the " & objService.Name & " service" & "" & "</font>"
    Call HTASleep(0)
    errReturn = objService.StartService()
    Call HTASleep(10)
    display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>Completed</font>"
    Logtype.innerHTML = "<font face=Arial font-size=11pt color=blue>" & "The Service " & objService.Name & " is now " & "Started on" & "</font>"
        
 End if  
  Next
restartC.disabled=false
stopServiceC.disabled=false
startServiceC.disabled=false
listServicesC.disabled=false
End Sub

'This part stops the service
'----------------------------------------------
Sub stopService
restartC.disabled=true
stopServiceC.disabled=true
startServiceC.disabled=true
listServicesC.disabled=true
strComputer=StoreEntry.Value
StoreName.value=UCase(strComputer)
Display.document.body.innerHTML=""
 
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    
Set colServiceList = objWMIService.ExecQuery _
    ("Select * from Win32_Service")

For Each objService in colServiceList
if objService.Name = serviceName.value then
	Logtype.innerHTML="Processing...."
	display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>" & "Stopping the " & objService.Name & " service " & "" & "</font>"
    Call HTASleep(0)
    errReturn = objService.StopService()
	Call HTASleep(10)
	display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>Completed</font>"
 	Logtype.innerHTML = "<font face=Arial font-size=11pt color=blue>" & "The Service " & objService.Name & " has been " & "Stopped" & "</font>"
    End if
    
  Next
restartC.disabled=false
stopServiceC.disabled=false
startServiceC.disabled=false
listServicesC.disabled=false
End Sub

'---------------------------------------------------
'This part restarts a service
'---------------------------------------------------
Sub restartService
restartC.disabled=true
stopServiceC.disabled=true
startServiceC.disabled=true
listServicesC.disabled=true
strComputer=StoreEntry.Value
StoreName.value=UCase(strComputer)
Display.document.body.innerHTML=""
 
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colServiceList = objWMIService.ExecQuery("Select * from Win32_Service")

For Each objService in colServiceList
if objService.Name = serviceName.value then
If objService.State = "Running" Then
	display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>"" Stopping the " & objService.Name & " Service on..""</font>"
	Logtype.innerHTML=" Stopping the " & objService.Name & " Service on.."  
    errReturn = objService.StopService()
    End if
End if

Next

Call HTASleep(10)

Set colServiceList = objWMIService.ExecQuery("Select * from Win32_Service")
 
For Each objService in colServiceList
  if objService.Name = serviceName.value then
  If objService.State = "Stopped" Then
  	 display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>"" Starting the " & objService.Name & " Service on..""</font>"
	 Logtype.innerHTML=" Starting the " & objService.Name & " Service on.."  
     errReturn = objService.StartService()
      Call HTASleep(5)
      display.document.body.InnerHTML = "<font face=Arial font-size=11pt color=green>Completed</font>"
 	  Logtype.innerHTML= objService.Name & " has been restarted on "
      End if
      End if
      Next
restartC.disabled=false
stopServiceC.disabled=false
startServiceC.disabled=false
listServicesC.disabled=false
End Sub
</SCRIPT>

<body>
<div style="border:0 solid #C0C0C0; position: absolute; width: 800px; height: 600px; z-index: 3; left: -2px; top: -22px; background-color:#C0C0C0" id="layer1">
&nbsp;</div>
<div style="position: absolute; width: 782px; height: 531px; z-index: 4; left: 5px; top: 39px; border-style: outset; border-width: 2px; padding: 1px" id="layer2">
&nbsp;</div>
<div style="position: absolute; width: 764px; height: 23px; z-index: 5; left: 14px; top: 532px; border-style: inset; border-width: 1px" id="layer3">
	<table border="1" width="100%" id="table1">
		<tr>
			<td width="123" style="font-family: Arial; font-size: 8pt; text-align: center">&nbsp;<span id="dateLabel" style="text-align: center; padding-left:5px"></span></td>
			<td height="20" width="22">&nbsp;</td>
			<td height="20" width="21">&nbsp;</td>
			<td height="20" width="418" style="font-size: 10pt; font-family: Arial">
			<p align="center">&nbsp;<span id="LogType">&nbsp;&nbsp; </span>
			<input id="StoreName" type="text" value="" name="T2" size="15" style="color: #008000; border-style: solid; border-width: 0; background-color: #C0C0C0; text-align:center; font-size:10pt; font-family:Arial"></td>
			<td height="20" width="22">&nbsp;</td>
			<td width="20">&nbsp;</td>
			<td width="82">
			<p align="center"><font face="Arial" style="font-size: 9pt">Version 
			1.0.0</font></td>
		</tr>
	</table>
</div>
<div style="position: absolute; width: 159px; height: 28px; z-index: 10; left: 404px; top: 538px; visibility:hidden" id="layer4">
	<iframe name="I1" id="Bar" src="CommCheckP.html" width="220" height="20" scrolling="no" border="0" frameborder="0">
	Your browser does not support inline frames or is currently configured not to display inline frames.
	</iframe></div>
<div style="border-style:inset; border-width:1px; position: absolute; width: 758px; height: 380px; z-index: 7; left: 16px; top: 149px" id="layer5">
<iframe name="I2" id="display" src="Display.html" width="755" height="375" border="0" frameborder="0" style="background-image: url('images/transparentblock.gif')">
Your browser does not support inline frames or is currently configured not to display inline frames.
</iframe></div>
<div style="position: absolute; width: 253px; height: 26px; z-index: 11; left: 12px; top: 12px" id="Controller">
	<table border="0" width="100%" id="table3" height="27">
		<tr>
			<td style="border-style: ridge; border-width: 1px" width="77" align="center">
			<font size="2" face="Arial">Start</font></td>
			<td style="border-style: ridge; border-width: 1px" width="82" align="center">
			<font face="Arial" size="2">Events Log</font></td>
			<td style="border-style: ridge; border-width: 1px" align="center">
			<font face="Arial" size="2">Services</font></td>
		</tr>
	</table>
</div>
<div style="position: absolute; width: 763px; height: 468px; z-index: 12; left: 14px; top: 60px; border-left-style: ridge; border-left-width: 1px; border-right-style: ridge; border-right-width: 1px; border-bottom-style: ridge; border-bottom-width: 1px" id="layer8">
&nbsp;</div>
<div style="position: absolute; width: 755px; height: 23px; z-index: 13; left: 18px; top: 46px" id="Entries">
	<table border="1" id="table4" height="28" width="755">
		<tr>
			<td width="156" style="border-style: ridge; border-width: 1px">
			<p align="right"><font size="2" face="Arial">&nbsp;&nbsp;System Name:</font></td>
			<td width="114" style="border-style: inset; border-width: 1px">
			<input id="StoreEntry" type="text" name="StoreEntry" size="18" style="border:1px ridge #C0C0C0; text-align:center; font-size:10pt; font-family:Arial; color:#008000"></td>
			<td width="105" style="border-style: ridge; border-width: 1px">
			<p align="right"><font face="Arial" size="2">Start Date:</font></td>
			<td width="114" style="border-style: inset; border-width: 1px">
			<input id="StartDate" type="text" value="StartDate" name="StartDate" size="18" style="border:1px ridge #C0C0C0; font-size:10pt; font-family:Arial; color:#008000; text-align:center"></td>
			<td width="108" style="border-style: ridge; border-width: 1px">
			<p align="right"><font face="Arial" size="2">End Date:</font></td>
			<td width="118">
			<input id="EndDate" type="text" value="EndDate" style="border:1px ridge #C0C0C0; text-align:center; font-size:10pt; font-family:Arial; color:#008000" size="18" tabindex="2" name="EndDate"></td>
		</tr>
	</table>
</div>
<div style="position: absolute; width: 284px; height: 36px; z-index: 14; left: 18px; top: 83px" id="layer9">
	<table border="1" width="101%" id="table5" style="border-style: solid; border-width: 0">
		<tr>
			<td style="border-style:solid; border-width:0px; " width="126">
			<b><font size="2" face="Arial" color="#008000">&nbsp;<u>EVENT QUERY</u></font></b></td>
			<td style="border-style:solid; border-width:0px; " width="204">
			<input id=runbutton  class="button" type="button" value=" Query Logs " name="run_button"  onClick="Pick" style="border-style: ridge; border-width: 1px; float:left"></td>
		</tr>
		</table>
</div>
<div style="position: absolute; width: 260px; height: 29px; z-index: 15; left: 268px; top: 82px" id="Eventtypes">
	<table border="1" width="35%" id="table6" style="border-style: ridge; border-width: 0">
		<tr>
			<td width="64" style="border-style: solid; border-width: 0px; ">
			<font face="Arial" size="2">Application</font></td>
			<td width="20"><input type="checkbox" name="ApplicationCheck" id="application" value="ApplicationEvents"></td>
			<td width="50" style="border-style: solid; border-width: 0px; ">
			<font size="2" face="Arial"> 
			Security</font></td>
			<td width="20"><font face="Arial"><input type="checkbox" id="Security" name="SecurityCheck" value="SecurityEvents"></font></td>
			<td width="47" style="border-style: solid; border-width: 0px; ">
			<font face="Arial" size="2">System</font></td>
			<td><input type="checkbox" name="SystemCheck" value="SystemEvents" id="System"></td>
		</tr>
		</table>
</div>
<div style="position: absolute; width: 755px; height: 30px; z-index: 16; left: 18px; top: 116px; border-style: solid; border-width: 0" id="layer11">
	<table border="1" width="97%" id="table7" style="border-style: solid; border-width: 0">
		<tr>
			<td style="border-style: solid; border-width: 0px"><b>
			<font face="Arial" size="2" color="#008000">SERVICE QUERY</font></b></td>
			<td style="border-style: solid; border-width: 0px">
			<p align="right"><font size="2" face="Arial">Name:</font></td>
			<td style="border-style: inset; border-width: 1px" width="142">
			<input id="serviceName" type=text style="color: #008000; font-size: 10pt; font-family: Arial; text-align: center; border: 1px ridge #C0C0C0; background-color: #FFFFFF" name="Servicename" size="18"></td>
			<td style="border-style: solid; border-width: 0px; " width="326">
			<input id="listServicesC" value="Query Services" type=button style="float: left; border-style: ridge; border-width: 1px" onclick="listServices" name="listservices"><input id="startServiceC" type="button" value="  Start " name="ServiceStart" style="border-style:outset; border-width:1px; float: left" onclick="startService"><p>
			<input id="stopServiceC" type="button" value="Stop" name="ServiceStop" style="float: left; border-style: ridge; border-width: 1px" onclick="stopService"></p>
			<input id="restartC" type=button value="Restart" name="restart" style="float: left; border-style: ridge; border-width: 1px" onclick="restartService"></td>
		</tr>
		</table>
</div>
<div style="position: absolute; width: 502px; height: 29px; z-index: 15; left: 267px; top: 106px; visibility:hidden" id="pathcredentials">
	<table border="1" width="82%" id="table8" style="border-style: ridge; border-width: 0">
		<tr>
			<td width="64" style="border-style: solid; border-width: 0px; ">
			<font face="Arial" size="2" color="#008000">Destination</font></td>
			<td width="195">
			<input id="destination" type=text style="font-size: 10pt; font-family: Arial; color: #008000; border-style: ridge; border-width: 1px; text-align:center" name="destination" size="30" value="C:\Eventlogs\"></td>
			<td width="75" style="border-style: solid; border-width: 0px; ">
			<p align="right">
			<font face="Arial" size="2" color="#008000">Extension</font></td>
			<td style="border-style: solid; border-width: 0">
			<p align="left">&nbsp;<input id="ext" type=text name="extensionT" style="font-size: 10pt; font-family: Arial; color: #008000; border-style: ridge; border-width: 1px; text-align:center" size="20" value=".CSV"></td>
		</tr>
		</table>
</div>
<div style="position: absolute; width: 730px; height: 17px; z-index: 17; left: 28px; top: 105px" id="layer12">
	<hr style="border-style: inset; border-width: 1px"></div>
</body>
'~~[/script]~~