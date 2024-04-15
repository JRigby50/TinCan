strComputer = "rsmvai-prn01"
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colLoggedEvents = objWMIService.ExecQuery _
        ("Select * from Win32_NTLogEvent Where Logfile = 'System' and " _
            & "EventCode = '7031'")

Wscript.Echo "Print Spooler Shutdown: " & colLoggedEvents.Count
	