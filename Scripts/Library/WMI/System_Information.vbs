'Uses WMI to retrieve the same data found in the System Information applet.

strComputer = "RSEMD4-BACKUP01"
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colSettings 
    Wscript.Echo "OS Name: " & objOperatingSystem.Name
    Wscript.Echo "Version: " & objOperatingSystem.Version
    Wscript.Echo "Service Pack: " & objOperatingSystem.ServicePackMajorVersion & "." & objOperatingSystem.ServicePackMinorVersion
    Wscript.Echo "OS Manufacturer: " & objOperatingSystem.Manufacturer
    Wscript.Echo "Windows Directory: " & objOperatingSystem.WindowsDirectory
    Wscript.Echo "Locale: " & objOperatingSystem.Locale
    Wscript.Echo "Available Physical Memory: " & objOperatingSystem.FreePhysicalMemory
    Wscript.Echo "Total Virtual Memory: " & objOperatingSystem.TotalVirtualMemorySize
    Wscript.Echo "Available Virtual Memory: " & objOperatingSystem.FreeVirtualMemory
    Wscript.Echo "OS Name: " & objOperatingSystem.SizeStoredInPagingFiles
Next
Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
For Each objComputer in colSettings 
    Wscript.Echo "System Name: " & objComputer.Name
    Wscript.Echo "System Manufacturer: " & objComputer.Manufacturer
    Wscript.Echo "System Model: " & objComputer.Model
    strTZ = (objComputer.CurrentTimeZone)/ 60
    Wscript.Echo "Time Zone: " & strTZ 'objComputer.CurrentTimeZone
    Wscript.Echo "Total Physical Memory: " & objComputer.TotalPhysicalMemory
Next
Set colSettings = objWMIService.ExecQuery("Select * from Win32_Processor")
For Each objProcessor in colSettings 
    Wscript.Echo "System Type: " & objProcessor.Architecture
    Wscript.Echo "Processor: " & objProcessor.Description
Next
Set colSettings = objWMIService.ExecQuery("Select * from Win32_BIOS")
For Each objBIOS in colSettings 
    Wscript.Echo "BIOS Version: " & objBIOS.Version
Next