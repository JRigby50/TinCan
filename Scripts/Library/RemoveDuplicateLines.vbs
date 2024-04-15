Const ForReading = 1, ForWriting = 2
Dim i, j 

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objInputFile = objFSO.OpenTextFile(".\input..txt", ForReading)
Set objOutputFile = objFSO.OpenTextFile (WorkingDir & ".\Output.txt", ForWriting, True)
Set objDict = CreateObject("Scripting.Dictionary")
j = 0 
'On Error Resume Next

While Not objInputFile.AtEndOfStream
	arrinputRecord = split(objInputFile.Readline, ",") 
 'Read input file with server names (Split is for later exspantion)
	strFirstField = arrinputRecord(0)                       
     
    WScript.Echo("Record: " & strFirstField) 
    
    if objDict.Exists(strFirstField) then
    	j=j+1
    Else
    	objDict.add strFirstField, strFirstField
    End if
Wend

'Output cleened list to a file

colKeys = objDict.Keys
For Each strKey in colKeys
    'wscript.echo "Result: " & strKey, objDict.Item(strKey)
    objOutputFile.writeline objDict.Item(strKey)
Next


'Report results
wscript.Echo "Total Records Writen:   " & objDict.count
wscript.Echo "Total Duplicates found: " & j

objInputFile.Close
objOutputFile.Close
