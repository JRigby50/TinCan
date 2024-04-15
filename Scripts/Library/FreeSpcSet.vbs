On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20
Const MByte = 1000000

arrComputers = Array("Cardiology2")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDisk", "WQL", _
                                      wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "FreeSpace: " & (objItem.FreeSpace/MByte) & " MB"
      WScript.Echo "Size: " & (objItem.Size/MByte) & " MB"
      PerCent = objItem.FreeSpace/objItem.Size
      Wscript.Echo PerCent * 100
      WScript.Echo
   Next
Next

