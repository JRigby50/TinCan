Const ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile("C:\RunningServicesList.csv", _
   ForAppending, True)
objLogFile.Write("Display Name,Service Name,Service State,Start Mode,Process ID ")
objLogFile.WriteLine


strComputer = InputBox("What computer would you like to retrieve the running services information from?")
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colRunningServices = objWMIService.ExecQuery _
    ("Select * from Win32_Service Where State = 'Running'")
For Each objService in colRunningServices
   objLogFile.Write(objService.DisplayName) & ","
   objLogFIle.Write(objService.Name) & ","
   objLogFIle.Write(objService.State) & ","
   objLogFIle.Write(objService.StartMode) & ","
   objLogFile.Write(objService.ProcessID) & ","
   objLogFile.WriteLine
    
Next

objLogFile.Close

