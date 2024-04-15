On Error Resume Next
CompName = inputbox("Enter name or IP address of local or remote computer", _ 
    "BIOSinfo")
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & CompName & "\root\cimv2")
If Err.Number <> 0 Then
    Wscript.Echo Err.Description
    Err.Clear
Else
Set colBIOS = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_BIOS")
For Each objBIOS in colBIOS
    Wscript.Echo "Computer name: " & CompName  & VbCrLF & VbCrLF & _
        "Manufacturer: " & objBIOS.Manufacturer & VbCrLF & _
        "Serial number: " & objBIOS.SerialNumber & VbCrLF & _
        "BIOS Version: " & objBIOS.SMBIOSBIOSVersion
Next
End If

