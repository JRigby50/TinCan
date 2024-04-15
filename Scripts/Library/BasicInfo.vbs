'Echoes the following information for a computer: computer name; operating system version; operating system build number; service pack number; amount of free memory; and location of the Windows folder. 

strcomputer = "."
'strcomputer = inputbox ("Please Type the Name of the Computer You Would Like to Get Info On","Get Machine Info","127.0.0.1")
	 For Each OS in GetObject("winmgmts:\\" & strcomputer).InstancesOf("Win32_OperatingSystem")
     Wscript.echo "Machine Info For " & os.csname & vbcrlf &_
     vbcrlf &_
     "PC Name: " & os.CSName & VBCRLF & _
     "OS: " & os.Caption & VBCRLF & _
     "Build= " & os.BuildNumber, os.BuildType & VBCRLF & _
     "Service Pack= " & os.CSDVersion & VBCRLF &_
     "Free Memory= " & os.FreePhysicalMemory & VBCRLF &_     
     "Windows Directory= " & os.WindowsDirectory
Next

