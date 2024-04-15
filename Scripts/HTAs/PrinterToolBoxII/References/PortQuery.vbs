On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array("LV48PCE00081501")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_TCPIPPrinterPort", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
'      WScript.Echo "ByteCount: " & objItem.ByteCount
'      WScript.Echo "Caption: " & objItem.Caption
'      WScript.Echo "CreationClassName: " & objItem.CreationClassName
'      WScript.Echo "Description: " & objItem.Description
'      WScript.Echo "HostAddress: " & objItem.HostAddress
'      WScript.Echo "InstallDate: " & WMIDateStringToDate(objItem.InstallDate)
'      WScript.Echo "Name: " & objItem.Name
'      WScript.Echo "PortNumber: " & objItem.PortNumber
'      WScript.Echo "Protocol: " & objItem.Protocol
'      WScript.Echo "Queue: " & objItem.Queue
'      WScript.Echo "SNMPCommunity: " & objItem.SNMPCommunity
'      WScript.Echo "SNMPDevIndex: " & objItem.SNMPDevIndex
'      WScript.Echo "SNMPEnabled: " & objItem.SNMPEnabled
'      WScript.Echo "Status: " & objItem.Status
'      WScript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
'      WScript.Echo "SystemName: " & objItem.SystemName
'      WScript.Echo "Type: " & objItem.Type
'      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function
