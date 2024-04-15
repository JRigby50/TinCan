Option Explicit
Dim strServer, strTest
strServer = "."

strTest = CPU_Architecture(strServer)
Wscript.Echo strTest

Function CPU_Architecture(strComputer)
'Determines the processor architecture (such as x86 or ia64) for a specified computer
Dim objWMIService, objProcessor, strArchitecture
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objProcessor = objWMIService.Get("win32_Processor='CPU0'")
	If objProcessor.Architecture = 0 Then
	    strArchitecture =  "x86"
	ElseIf objProcessor.Architecture = 1 Then
	    strArchitecture =  "MIPS"
	ElseIf objProcessor.Architecture = 2 Then
	    strArchitecture =  "Alpha"
	ElseIf objProcessor.Architecture = 3 Then
	    strArchitecture =  "PowerPC"
	ElseIf objProcessor.Architecture = 6 Then
	    strArchitecture =  "ia64"	
	Else
	    strArchitecture =  "Unknown"
	End If
CPU_Architecture = strArchitecture
End Function