strComputer = "RSEMD1-CITRIX02"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile(strComputer & ".txt", True)

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colQuickFixes = objWMIService.ExecQuery _
    ("Select * from Win32_QuickFixEngineering")

For Each objQuickFix in colQuickFixes
    objTextFile.WriteLine "Computer: " & objQuickFix.CSName
    objTextFile.WriteLine "Description: " & objQuickFix.Description
    objTextFile.WriteLine "Hot Fix ID: " & objQuickFix.HotFixID
    objTextFile.WriteLine "Installation Date: " & objQuickFix.InstallDate
    objTextFile.WriteLine "Installed By: " & objQuickFix.InstalledBy
    objTextFile.WriteLine
    objTextFile.WriteLine
Next
objTextFile.Close