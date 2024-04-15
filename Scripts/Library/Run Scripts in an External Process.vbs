'Demonstrates how to use an environment variable, formatted like the PATH variable to include VBScript code,
'and then execute scripts in an external process , all without having to have a fixed location for the VBScript code.
'This makes central storage of scripts (repository-style) possible. 


' Possible usage
' Include the external Class library script (to allow logging to a file from scripts)
Call IncludeFile("Lib_ClassLogging.vbs")   
 
' Execute the external VBScript (to send an email to the current system Standby)
Call intRunScriptByPath("Alert_StandBy.vbs")
  
Private Sub IncludeFile(strFileName)
  Const ForReading = 1
  Dim oFS
  Dim strFullFileName
  Set oFS = WScript.CreateObject("Scripting.FileSystemObject")
  ' In this sample, %PATH% is used as the environment variable to search
  ' You can create a (process) environment variable specifically to identify the Script library/libraries, 
  ' e.g. %VBSLIB%, which can point to a central share in your network, like "\\Server1\VBScriptReporitory"
  strFullFileName = FindFileByEnvironment(strFileName, "%PATH%") 
  If strFullFileName <> "" Then 
    ' The filename has been found somewhere
    WScript.Echo "Including: " & strFullFileName   
    Dim f: Set f = oFS.OpenTextFile(strFullFileName, ForReading)
    Dim s: s = f.ReadAll
    f.Close
    WScript.ExecuteGlobal s
  Else
    'Process the FileNotFound error
    DisplayError "*** FATAL ERROR: Include file " & FileName & " not found."
    WScript.Quit
  End If
End Sub

Private Function intRunScriptByPath(strCommand, strParameters, intWindowStyle, bWaitOnReturn)
  Dim strFullCommand
  Dim oWS
  Set oWS = WScript.CreateObject("WScript.Shell")
  intRunScriptByPath = -1
  strFullCommand = FindFileByEnvironment(strCommand, "%PATH%")
  If strFullCommand <> "" Then
    WScript.Echo "Running: " & strFullCommand
    intRunScriptByPath = oWS.Run("cscript.exe //nologo " & strFullCommand, intWindowStyle, bWaitOnReturn)
  Else
    'Process the FileNotFound error
    DisplayError "*** FATAL ERROR: External script file "  & strCommand 
  End If
End Function

Private Sub DisplayError(ErrorText)
  MsgBox ErrorText, 0, Wscript.ScriptName
End Sub

Private Function strFindFileByEnvironment(strFileName, strVariableName)
  const vbTextCompare = 1
  Dim oFS, oWS   
  Dim strFullFileName, arr_strPath, iCount
  Set oFS = WScript.CreateObject("Scripting.FileSystemObject")
  Set oWS = WScript.CreateObject("WScript.Shell")
  ' first, check the location where the current script was started from
  strFullFileName = oFS.BuildPath(oFS.GetParentFolderName(WScript.ScriptFullName), strFileName)
  If oFS.FileExists(strFullFilename) Then
    strFindFileByPath = strFullFileName
    Exit Function
  End If
  ' The file was not found in the script's home folder, so we search all elements in the environment variable
  ' The first one to match is returned to the caller
  arr_strPath = split(oWS.ExpandEnvironmentStrings(strVariableName), ";", -1, vbTextCompare)
  If Not IsEmpty(arrPath) Then
    For iCount = 0 To uBound(arr_strPath)
      strFullFileName = oFS.BuildPath(arr_strPath(i), strFileName)
      If oFS.FileExists(strFullFileName) Then 
        strFindFileByPath = strFullFileName
        Exit Function
      End If
    Next
  End If
  ' Nothing was found, so we return an empty string
  strFindFileByPath = ""
End function
