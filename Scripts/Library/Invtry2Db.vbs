Retrieves inventory information from multiple computers and then writes that information to a pre-created database. 

Const CONVERSION_FACTOR = 1048576
On Error Resume Next
Set objADSysInfo = CreateObject("ADSystemInfo")
strUser = objADSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
strComputer = "."
strCompName = objADSysInfo.ComputerName
Set objComputerName = GetObject("LDAP://"& strCompName)
objComputerName.GetInfo
strComputerName = objComputerName.Get ("Name")
objUser.GetInfo
strdisplayName = objUser.Get ("displayName")
'Msgbox strdisplayName, 0, strMessage
strdescription = objUser.Get ("description")
strdescription = mid(strdescription,5,len(strdescription) -4)
'** Opens the Data Source "Inventory" and writes to the appropriate table

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordset = CreateObject("ADODB.Recordset")
objConnection.Open "DSN=Inventory;"
objRecordset.CursorLocation = adUseClient
objRecordset.Open "SELECT * FROM Hardware" , objConnection, _
adOpenStatic, adLockOptimistic
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer& "\root\cimv2")
Set colSoundCards = objWMIService.ExecQuery _
("SELECT * FROM Win32_SoundDevice")
Set colItems = objWMIService.ExecQuery _
("SELECT * FROM Win32_Processor")
Set colBoards = objWMIService.ExecQuery _
("SELECT * FROM Win32_BaseBoard")
Set colCDROMs = objWMIService.ExecQuery _
("SELECT * FROM Win32_CDROMDrive")
Set colDrives = objWMIService.ExecQuery _
("SELECT * FROM Win32_DiskDrive")
Set colOSes = objWMIService.ExecQuery _
("SELECT * FROM Win32_OperatingSystem")
Set colVideoCards = objWMIService.ExecQuery _
("SELECT * FROM Win32_VideoController")
Set colSystems = objWMIService.ExecQuery _
("SELECT * FROM Win32_ComputerSystem")
Set colApps = objWMIService.ExecQuery _
("Select * from Win32_Product Where Caption Like '%Microsoft Office%'")

strSearchCriteria = "ComputerName = '" & strComputerName &"'"
'Wscript.Echo strSearchCriteria
objRecordSet.Find strSearchCriteria

If objRecordSet.EOF Then
'Wscript.Echo "Record Not Found"
For Each objSoundCard in colSoundCards
objRecordset.AddNew
objRecordset("ComputerName") = objSoundCard.SystemName
objRecordset("Sound Card") = objSoundCard.ProductName
objRecordset("User") = strdisplayName
objRecordset("Department") = strdescription
objRecordset.Update
Next
For Each objItem in colItems
objRecordset("ProcessorName") = objItem.Name
objRecordset("ClockSpeed") = objItem.MaxClockSpeed
objRecordset.Update
Next
For Each objBoard in colBoards
objRecordset("Motherboard Manufacturer") = objBoard.Manufacturer
objRecordset("Motherboard Model") = objBoard.Product
objRecordset.Update
Next
For Each objCDROM in colCDROMs
objRecordset("Optical Drive Type") = objCDROM.Name
objRecordset.Update
Next
For Each objDrive in colDrives
objRecordset("Hard Drive Model") = objDrive.Model
objRecordset("Hard Drive Capacity") = Int(objDrive.Size / CONVERSION_FACTOR) & "MB"
objRecordset.Update
Next
For Each objVideoCard in colVideoCards
objRecordset("Video Card Name") = objVideoCard.Description
objRecordset("Video Card Memory") = Int(objVideoCard.AdapterRAM / CONVERSION_FACTOR) & "MB"
objRecordset.Update
Next
For Each objOSe in colOSes
objRecordset("Operating System") = objOSe.Caption
objRecordset.Update
Next
For Each objSystem in colSystems
objRecordset("Physical RAM") = Int(objSystem.TotalPhysicalMemory / CONVERSION_FACTOR) & "MB"
objRecordset.Update
Next
For Each objApp in colApps
objRecordset("Office Version") = objApp.Caption
objRecordset.Update
Next
objRecordset.Close
objConnection.Close

Else
'Wscript.Echo "Record Found"
For Each objSoundCard in colSoundCards
objRecordset("ComputerName") = objSoundCard.SystemName
objRecordset("Sound Card") = objSoundCard.ProductName
objRecordset("User") = strdisplayName
objRecordset("Department") = strdescription
objRecordset.Update
Next
For Each objItem in colItems
objRecordset("ProcessorName") = objItem.Name
objRecordset("ClockSpeed") = objItem.MaxClockSpeed
objRecordset.Update
Next
For Each objBoard in colBoards
objRecordset("Motherboard Manufacturer") = objBoard.Manufacturer
objRecordset("Motherboard Model") = objBoard.Product
objRecordset.Update
Next
For Each objCDROM in colCDROMs
objRecordset("Optical Drive Type") = objCDROM.Name
objRecordset.Update
Next
For Each objDrive in colDrives
objRecordset("Hard Drive Model") = objDrive.Model
objRecordset("Hard Drive Capacity") = Int(objDrive.Size / CONVERSION_FACTOR) & "MB"
objRecordset.Update
Next
For Each objOSe in colOSes
objRecordset("Operating System") = objOSe.Caption
objRecordset.Update
Next
objRecordset.Close
objConnection.Close
End If

