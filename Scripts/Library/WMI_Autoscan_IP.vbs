Option Explicit
Dim dtmThisDay, dtmThisMonth, dtmThisYear
Dim strServer, strTotalMemory, strMemory, strCPUName, strOpSys, strOpVer, strOS, strOSVer, strCPU, strCPUCount, strIPAddresses, strSQL
Dim strSite, strRecordID, strOS_Name, strOS_Version, strRAM, strCPU_Count, strCPU_Description, strUpdateSQL, objScan
Dim strRaw, strClean, strTmp1, arrTmp
Dim objFSO, objServerList, objLog, objWMIService, objPing, objStatus, objItem, colItems, objErrorsOut, objConn, objServers, objUpdate, objOldIPAddresses
Dim bolServerOK, bolUpdate, varTotalMemory, bCPUCount, cpuCounter, s, u, bolUpdateDB, bolUpdateIP, objOldIPAddress
Dim objDict, strItem, arrTemp, arrNew, j, i, strTmp, arrIPAddresses, arrIPAddressesWMI, z

Dim strOldIPAddress, arrOldIPAddresses, dicIP, dicOldIP, colKeys, strKey

'Initialize index and set Field Names for arrUpdate
Dim arrUpdate()
s=1
Const Site = 0
Const Server = 1
Const RecordID = 2
Const OS_Name = 3
Const OS_Version = 4
Const RAM = 5
Const CPU_Count = 6
Const CPU_Description = 7

Const ForReading = 1
Const ForWriting = 2

dtmThisDay = Day(Date)
dtmThisMonth = Month(Date)
dtmThisYear = Year(Date)
bolUpdateDB=False
bolUpdateIP=False
'*****BEGIN MAIN ROUTINE*****
' Open Log and Error Log
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLog = objFSO.CreateTextFile("c:\AdminReports\Log.csv", ForWriting)
Set objErrorsOut = objFSO.CreateTextFile("c:\AdminReports\Errors.txt", ForWriting)
'Open the DB Connection
Set objConn = CreateObject("ADODB.Connection")
objConn.Open "DSN=NGServers;Uid=Northgrum\yyMD1-ESWinTelDB;Pwd=Geronimo1829;"


'*****Get IPaddresses to put into TServerSub table
strSQL = "SELECT Site, Server, RecordID FROM TServer WHERE WMI_Autoscan = '1'"
Set objScan = objConn.Execute(strSQL)
Do Until objScan.EOF
' Get current IP Addresses from Server
	strSite = objScan("Site")
	strServer = objScan("Server")
	strRecordID = objScan("RecordID")
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strServer & "\root\cimv2")
	On Error Resume Next
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration",,48) 
	strTmp = ""
	Set dicIP = CreateObject("Scripting.Dictionary")
	For Each objItem in colItems 
	    If isNull(objItem.IPAddress) Then
	        'Do nothing
	    Else
			If Not dicIP.Exists(strItem) Then
			strTmp1 = Join(objItem.IPAddress, ",")
	        dicIP.Add strTmp1, strTmp1  
	        End If  
	    End If
	Next

'*****Get Old IP Addresses from DB
	
	strSQL = "SELECT Value, RecordID FROM TServerSub WHERE Site = '" & strSite & "' AND Server = '" & strServer & "' AND RecordType = 'IPAddress'"
	Set objOldIPAddresses = objConn.Execute(strSQL)
	Set dicOldIP = CreateObject("Scripting.Dictionary")

	strTmp = ""
	Do Until objOldIPAddresses.EOF
	strOldIPAddress= objOldIPAddresses(0)
	strRecordID = objOldIPAddresses(1)
		If strOldIPAddress = "" Then
		'Nothing
		Else
			If Not dicOldIP.Exists(strOldIPAddress) Then  
				dicOldIP.Add strOldIPAddress, strRecordID  
            End If  
		End If
	objOldIPAddresses.MoveNext
	Loop

' Do any IPs need to be removed?
colKeys = dicOldIP.Keys
For Each strKey in colKeys
	If dicIP.Exists(strKey) Then 
	'Do Nothing
	Else
	'Delete from the DB
	strSQL = "DELETE FROM TServerSub WHERE (Site = '" & strSite & "' AND Server = '" & strServer & "' AND RecordID = '" & strRecordID &  "' AND RecordType = 'IPAddress' AND Value = '" & oldIPAddress & "')"
	objLog.WriteLine strSQL
	End If
'updateCommands.append(strSQL)
Next

' Do any IPs need to be added?
colKeys = dicIP.Keys
For Each strKey in colKeys
	If dicOldIP.Exists(strKey) Then 
	'Do Nothing
	Else
	'Add to the DB
	strSQL = "INSERT INTO TServerSub(Site,Server,RecordType,Value) VALUES ('" & strSite & "','" & strServer & "','IPAddress','" & strIPAddress & "')"
	objLog.WriteLine strSQL
	End If
'updateCommands.append(strSQL)


Next


	    
'	        ReDim arrTemp(objDict.Count - 1)  
'	        arrNew = objDict.Items
'        For i = UBound(arrNew) - 1 To 0 Step -1  
'            For j = 0 To i  
'                If arrNew(j)>arrNew(j+1) Then  
'                strTmp = arrNew(j+1)  
'                arrNew(j+1) = arrNew(j)  
'                arrNew(j) = strTmp  
'                End If  
'            Next  
'        Next 

	
' Compare arrIPAddresses to arrOldIPAddresses 
'	For Each oldIPAddress in arrOldIPAddresses
'		If OldIPAddress not in arrIPAddresses Then
'		strSQL = "DELETE FROM TServerSub WHERE (Site = '" + Site + "' AND Server = '" + Server
'		strSQL = "' AND RecordType = 'IPAddress' AND Value = '" & oldIPAddress & "')"
'		updateCommands.append(strSQL)
'		End If
'	Next
'	For Each IPAddress in arrIPAddresses:
'		If (IPAddress not in oldIPAddresses Then
'		strSQL = "INSERT INTO TServerSub(Site,Server,RecordType,Value) VALUES ('"
'		strSQL += Site & "','" + Server & "','IPAddress','" + IPAddress & "')"
'		updateCommands.append(strSQL)
'		End If
'	Next
'Next
'If bolUpdateIP=True Then
' Get IPaddresses and put into TServerSub table
' strUpdateSQL = "Update TServerSub
'Set objUpdate = objConn.Execute(strUpdateSQL)
'End If
'Set objUpdate = objConn.Execute(strUpdateSQL)

'*******************************************************************************************
'*******************************************************************************************
	'End If
	WScript.Echo (strServer)
	objScan.MoveNext
Loop

'*****Cleanup to End Gracefully
objConn.Close
objErrorsOut.Close
objLog.Close
'*****END MAIN ROUTINE*****
'*******************************************************************************************
'*******************************************************************************************
'*****FUNCTIONS START*****
'*****Get OS
Function OS
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strServer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48)
		For Each objItem in colItems
		    strOpSys = objItem.Caption
		Next
	OS = strOpSys
End Function
'*****Get OS Version
Function OSVer
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strServer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48)
		For Each objItem in colItems
		    strOpVer = objItem.CSDVersion
		Next
	OSVer = strOpVer
End Function

Function CPU
'CPU Description
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strServer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor",,48)
	    For Each objItem in colItems
			strCPUName = objItem.Name
		    Next
	CPU = strCPUName
End Function

Function CPUCount
'CPU Count
	cpuCounter = 0
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strServer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor",,48)
	    For Each objItem in colItems
			strCPUName = objItem.Name
		    cpuCounter = cpuCounter +1
		Next
	CPUCount = cpuCounter
End Function

Function Memory
'Get Total Memory
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strServer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PhysicalMemory",,48) 
varTotalMemory = 0.0
	For Each objItem in colItems 
		varTotalMemory = varTotalMemory + int(objItem.Capacity)
	Next
' Put Memory into proper format
varTotalMemory = varTotalMemory / 1024
	If varTotalMemory < 1024 Then
		strTotalMemory =(round(varTotalMemory)) & " KB"
		Memory=strTotalMemory
	Exit Function
	End If
varTotalMemory = varTotalMemory / 1024
		If varTotalMemory < 1024 Then
		strTotalMemory = (round(varTotalMemory)) & " MB"
		Memory=strTotalMemory
		Exit Function
		End If
varTotalMemory = varTotalMemory / 1024
		If varTotalMemory < 1024 Then
		strTotalMemory = (round(varTotalMemory)) & " GB"
		Memory=strTotalMemory
		Exit Function
        End If
varTotalMemory = varTotalMemory / 1024
		strTotalMemory = (round(varTotalMemory)) & " TB"
		Memory=strTotalMemory
End Function

Function CleanString(strRaw)
'Remove garbage from description strings
	strClean = Replace(strRaw,"Microsoft", "")
	strClean = Replace(strClean,"MICROSOFT", "")
	strClean = Replace(strClean,"(R)", "")
	strClean = Replace(strClean,"(r)", "")
	strClean = Replace(strClean,"(TM)", "")
	strClean = Replace(strClean,"(tm)", "")
	strClean = Replace(strClean,"®", "")
	strClean = Replace(strClean,",", " ")
	strClean = Trim (strClean)
		While Instr(strClean, "  ")
		strClean = Replace(strClean,"  ", " ")
		Wend
	CleanString = strClean
End Function

Function Sort(arrTemp)
'Takes an Array and Returns an Array sorted in Ascending order with duplicates removed.
Set objDict = CreateObject("Scripting.Dictionary")  
    For Each strItem In arrTemp  
            If Not objDict.Exists(strItem) Then  
                    objDict.Add strItem, strItem  
            End If  
    Next  
        ReDim arrTemp(objDict.Count - 1)  
        arrNew = objDict.Items  
        For i = UBound(arrNew) - 1 To 0 Step -1  
            For j = 0 To i  
                If arrNew(j)>arrNew(j+1) Then  
                        strTmp = arrNew(j+1)  
                        arrNew(j+1) = arrNew(j)  
                        arrNew(j) = strTmp  
                End If  
            Next  
        Next  
Sort = arrNew
End Function
'*****FUNCTIONS END*****
'