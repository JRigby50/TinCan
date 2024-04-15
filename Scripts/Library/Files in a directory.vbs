strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService. _
    ExecQuery("Select * from CIM_DataFile where Path = '\\ton convert\\'")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile("C:\Documents and Settings\rigbyja\Desktop\eBooks\ton convert\convert.txt", True)

For Each objFile in colFiles
strIn=objFile.Name & Chr(32)
strOut= strIn & "(1) "
objTextFile.WriteLine "pmrvHgwP.py " &  strIn &  strOut &   " GXGWKNV$9A"
Next