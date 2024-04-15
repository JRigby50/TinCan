strComputer = "."
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile("Scheduled Jobs.csv", True)
objTextFile.WriteLine "Caption" & "," & "Command" & "," & "DaysOfMonth" & "," & "DaysOfWeek" & "," &  _
    "Description" & "," & "ElapsedTime" & "," & "InstallDate" & "," & "InteractWithDesktop" & "," & "JobId" & "," & _
    "JobStatus" & "," & "Name" & "," & "Notify" & "," & "Owner" & "," & "Priority" & "," & "RunRepeatedly" & "," & "StartTime" & "," & _
    "Status" & "," & "TimeSubmitted" & "," & "UntilTime"
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_ScheduledJob",,48)
    
For Each objItem in colItems 
    objTextFile.Write objItem.Caption & ","
    objTextFile.Write objItem.Command & ","
    objTextFile.Write objItem.DaysOfMonth & ","
    objTextFile.Write objItem.DaysOfWeek & ","
    objTextFile.Write objItem.Description & ","
    objTextFile.Write objItem.ElapsedTime & ","
    objTextFile.Write objItem.InstallDate & ","
    objTextFile.Write objItem.InteractWithDesktop & ","
    objTextFile.Write objItem.JobId & ","
    objTextFile.Write objItem.JobStatus & ","
    objTextFile.Write objItem.Name & ","
    objTextFile.Write objItem.Notify & ","
    objTextFile.Write objItem.Owner & ","
    objTextFile.Write objItem.Priority & ","
    objTextFile.Write objItem.RunRepeatedly & ","
    objTextFile.Write objItem.StartTime & ","
    objTextFile.Write objItem.Status & ","
    objTextFile.Write objItem.TimeSubmitted & ","
    objTextFile.Write objItem.UntilTime
    objTextFile.WriteLine
Next
objTextFile.Close