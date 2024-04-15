Set objWMIService = GetObject("winmgmts:")
Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='c:'")
Wscript.Echo objLogicalDisk.FreeSpace