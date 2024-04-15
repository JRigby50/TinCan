'WMI Scan to SQL DB script to replace Python Script written by David Price
'This script Scans through ESWINTELDB.TServer and, for each server marked for WMI_Autoscan, verifies the data and performs any updates needed.
'This DB is accessed by http://eswintel.md.essd.northgrum.com
'Script Written by Jim Rigby
'james.rigby@ngc.com
'Version 1.5
'April 7, 2011
Option Explicit
Dim dtmThisDay, dtmThisMonth, dtmThisYear
Dim strServer, strTotalMemory, strMemory, strCPUName, strOpSys, strOpVer, strOS, strOSVer, strCPU, strCPUCount, strIPAddresses, strSQL
Dim strSite, strRecordID, strOS_Name, strOS_Version, strRAM, strCPU_Count, strCPU_Description, objScan
Dim strRaw, strClean, dicIP, dicOldIP, objShell
Dim objFSO, objServerList, objLog, objWMIService, objPing, objStatus, objItem, colItems, objConn, objServers, objUpdate, objOldIPAddresses
Dim bolServerOK, bolUpdate, varTotalMemory, cpuCounter, s, u, bolUpdateDB, bolUpdateIP
Dim objDict, strItem, arrTemp, arrNew, j, i, strTmp, arrIPAddresses, arrIPAddressesWMI, z
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
Set objLog = objFSO.CreateTextFile("Log.txt", ForWriting)
'Set objErrorsOut = objFSO.CreateTextFile("Errors.txt", ForWriting)
'Open the DB Connection
Set objConn = CreateObject("ADODB.Connection")
objConn.Open "DSN=NGServers;yyMD1-ESWinTelDB;Pwd=Geronimo1829;" 'Uid=Northgrum\rigbyja;Pwd=Katrina05" '"
'objConn.Open "Provider=sqloledb;Data Source=lv48pce00081501\MYDB;Initial Catalog=NGSERVERS;User Id=Northgrum\yyMD1-ESWinTelDB;Password=Geronimo1829;"
'Get all the servers With WMI Scan Set To 1
strSQL = "SELECT Site, Server, RecordID, OS_Name, OS_Version, RAM, CPU_Count, CPU_Description FROM TServer WHERE WMI_Autoscan = 1"
Set objServers = objConn.Execute(strSQL)
'Cycle through the servers looking for changes
Do Until objServers.EOF
	strSite = objServers("Site")
	strServer = objServers("Server")
	strRecordID = objServers("RecordID")
	strOS_Name = objServers("OS_Name")
	strOS_Version = objServers("OS_Version")
	strRAM = objServers("RAM")
	strCPU_Count = objServers("CPU_Count")
	strCPU_Description = objServers("CPU_Description")
	On Error Resume Next
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strServer & "\root\cimv2")
'Is the server online?
	bolServerOK = True
	Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
	ExecQuery("select * from Win32_PingStatus where address = '" & strServer & "'")
		For Each objStatus in objPing 
	        If IsNull(objStatus.StatusCode) Or objStatus.StatusCode<>0 Then
	            objLog.WriteLine "Server " & strServer & " is not reachable!"
	            bolServerOK = False
	        End If
		Next
	If bolServerOK = True Then
' Process Server
		strOS = OS
		strOS = CleanString(strOS)
		strOSVer = OSVer
		strCPU = CPU
		strCPU = CleanString(strCPU)
		strCPUCount =CPUCount
		strMemory = Memory
		bolUpdate = False
		
		If strOS_Name <> strOS Then
		bolUpdate = True
		End If
		
		If strOS_Version <> strOSVer Then
		bolUpdate = True
		End If
		
		If strRAM <> strMemory Then
		bolUpdate = True
		End If
		
		If strCPU_Description  <> strCPU Then
		bolUpdate = True
		End If
		
		If strCPU_Count <> strCPUCount Then
		bolUpdate = True
		End If
		
		If bolUpdate = True Then
		bolUpdateDB = True
		ReDim Preserve arrUpdate(7,s)
		objLog.WriteLine strServer & "," &  strOS & "," &  strOSVer & "," &  strCPU & "," & strCPUCount & "," &  strMemory
		arrUpdate(Site, s) = strSite
		arrUpdate(Server, s) = strServer
		arrUpdate(RecordID, s) = strRecordID
		arrUpdate(OS_Name, s) = strOS
		arrUpdate(OS_Version, s) = strOSVer
		arrUpdate(RAM, s) = strMemory
		arrUpdate(CPU_Count, s) = strCPUCount
		arrUpdate(CPU_Description, s) = strCPU
		s = s + 1
		End If
		End If
	WScript.Echo (strServer)
	objServers.MoveNext
Loop
'Write Updates to DB
If bolUpdateDB = True Then
	For u = 1 to s 'Each strUpdate In arrUpdate
			strServer = arrUpdate(Server,u)
			strOS = arrUpdate(OS_Name,u)
			strOSVer = arrUpdate(OS_Version,u)
			strCPU = arrUpdate(CPU_Description,u)
			strCPUCount = arrUpdate(CPU_Count,u)
			strMemory = arrUpdate(RAM,u)
	strSQL = "Update TServer Set OS_Name = '" & strOS & "',OS_Version = '" & strOSVer & "',CPU_Description = '" & strCPU & "',CPU_Count = '" & strCPUCount & "',RAM = '" & strMemory & "' Where Server = '" & strServer & "'"
	objLog.WriteLine strSQL
	'objConn.Execute(strSQL)
	Next
End If
'*******************************************************************************************
WScript.Echo ("Starting IP Section")
'*******************************************************************************************

'*****Get IPaddresses to put into TServerSub table
strSQL = "SELECT Site, Server, RecordID FROM TServer WHERE WMI_Autoscan = '1'"
Set objScan = objConn.Execute(strSQL)
Do Until objScan.EOF
' Get current IP Addresses from Server
	strSite = objScan("Site")
	strServer = objScan("Server")
	strRecordID = objScan("RecordID")
	'Is the server online?
	bolServerOK = True
	Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
	ExecQuery("select * from Win32_PingStatus where address = '" & strServer & "'")
		For Each objStatus in objPing 
	        If IsNull(objStatus.StatusCode) Or objStatus.StatusCode<>0 Then
	            objLog.WriteLine "Server " & strServer & " is not reachable!"
	            bolServerOK = False
	        End If
		Next
If bolServerOK = True Then
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
	'objConn.Execute(strSQL)
	End If
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
	Next
	WScript.Echo (strServer)
	'objConn.Execute(strSQL)
End if
	objScan.MoveNext
Loop





'*****Cleanup to End Gracefully
objConn.Close
'objErrorsOut.Close
objLog.Close
'Send email

Set objShell= CreateObject("WScript.Shell")
objShell.Run "C:\bin\Blat.exe Log.txt -from netsysnotif@ngc.com -server relaymd001.northgrum.com -to james.rigby@ngc.com -subject ""WMI Autoscan"" -html",0,True
'*****END MAIN ROUTINE*****

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
                    objDict.Add sItem, sItem  
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
