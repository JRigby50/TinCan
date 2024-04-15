'**************************************************************************
'* FolderCleaner.vbs
'*
'* This script will delete all files older than a specified number of days
'* from the specified folder.
'*
'* It accepts no arguments and outputs a logfile C:\ReportCleaner.log
'*
'* Created 29 September 2003 by Chris Bennell
'*
'**************************************************************************


strfolder = "\\RSMV48-FPS01\FPSECURITY" 'The folder to delete files from
intAgeToDelete = 4 ' The age (in days) to delete files if equal or older.

intDeletedCount = 0
LogFilename = "C:\ReportCleaner.log"
Set fso = CreateObject("Scripting.FileSystemObject")
Set OutputFile = FSO.CreateTextFile(LogFilename,True)
Set WshNetwork = WScript.CreateObject("WScript.Network") 'declaration of a network object for storing computer name
Set checkfolder = fso.GetFolder(strfolder)

computer = WshNetwork.ComputerName

OutputFile.write WScript.ScriptName & " script run on " & computer & ", date: " & NOW()
OutputFile.write vbCrlf & "This script will delete all files equal to or older than " & intAgeToDelete & " days in the " & strfolder & " folder."
CheckFiles(checkfolder) 'Check all files in the folder and delete old ones.
OutputFile.write vbCrlf & intDeletedCount & " files deleted."
OutputFile.write vbCrlf & "********** End of ReportCleaner log **********"
OutputFile.Close
Set fileObject = Nothing
Set fso = Nothing

'----------------------------
Sub CheckFiles(Folder)
'This procedure simply passes all files in a folder to the DeleteFiles procedure
    For Each File In Folder.Files
      Checkstring = File.Path
      DeleteFiles(Checkstring)
      If err.number <> 0 Then OutputFile.Write vbCrlf & err.description
    Next
End Sub
'----------------------------

'----------------------------
Sub DeleteFiles(strfile)
'This procedure checks the date a file was created and deletes it if is older than a specfied number of days
  Set fileObject = fso.GetFile(strfile)
'  wscript.echo fileObject.Name
'  wscript.echo fileObject.DateCreated

  DaysOld = Round(now() - fileObject.DateLastModified)
  wscript.echo fileObject.Name & " created " & fileObject.DateLastModified & " is " & DaysOld & " days old."
  If DaysOld >= intAgeToDelete Then   
    OutputFile.write vbCrlf & fileobject.name & " was created earlier than the specified number of days"
'   Wscript.Echo "Deleted "& fileObject.path & fileObject.name
    OutputFile.write vbCrlf & "Deleted "& fileObject
    fso.DeleteFile(fileObject)
    intDeletedCount = intDeletedCount +1
  End If
End Sub
'----------------------------

