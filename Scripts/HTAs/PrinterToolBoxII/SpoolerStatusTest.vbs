strServer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strServer & "\root\cimv2")
Set colRunningServices =  objWMIService.ExecQuery ("Select * from Win32_Service Where Name = 'Spooler'")

For Each objService in colRunningServices 
    strSpoolerStatus = objService.State
    MsgBox strSpoolerStatus
'    If strSpoolerStatus = "Running" Then
'    SpoolerStatus.InnerHTML=strOK & strSpoolerStatus
'    Else
'    SpoolerStatus.InnerHTML=strWarning & strSpoolerStatus
'    End If
Next

<font color="black" face="arial" size="1"> 1091

<font color="black" face="arial" size="2"><B>Status: </B><span name="HTAStatus" id="HTAStatus" rows="1" cols="110" title="Displays the current status of the program"></span><BR><BR>

<font color="black" face="arial" size="2"><B>Spooler: </B><span name="SpoolerStatus" id="SpoolerStatus" rows="1" cols="110" title="Displays the current status of the Spooler Service"></span><BR><BR>
