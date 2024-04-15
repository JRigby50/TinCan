'determine whether a computer is a tower, a mini-tower, a laptop, and so on?

'Use the Win32_SystemEnclosure class and check the value of the ChassisType property.
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colChassis = objWMIService.ExecQuery ("Select * from Win32_SystemEnclosure")
For Each objChassis in colChassis
    For Each objItem in objChassis.ChassisTypes
        Wscript.Echo "Chassis Type: " & objItem
    Next
Next