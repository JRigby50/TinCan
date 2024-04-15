Option Explicit
Dim strServer, arrTest
strServer = "."

arrTest = SN_TN(strServer)


Function SN_TN(strComputer)
'Get the serial number and asset tag of a computer
Dim objWMIService, colSMBIOS, objSMBIOS, arrSN_TN(), x
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSMBIOS = objWMIService.ExecQuery ("Select * from Win32_SystemEnclosure")
x = 0
For Each objSMBIOS in colSMBIOS
	ReDim Preserve arrSN_TN(2,x)
    arrSN_TN(0,x) = "Part Number: " & objSMBIOS.PartNumber
    arrSN_TN(1,x) = "Serial Number: " & objSMBIOS.SerialNumber
    arrSN_TN(2,x) = "Asset Tag: " & objSMBIOS.SMBIOSAssetTag
    x = x + 1
Next
SN_TN = arrSN_TN
End Function