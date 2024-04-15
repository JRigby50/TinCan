Option Explicit

Const programName = "autg.vbs"
Const verNum = "1.1"

' Globals
Dim group          ' Short name of group
Dim groupDN        ' Distinguished name of group

Dim userFSO        ' File system object to read user file
Dim userFileName   ' User file name
Dim objUserFile    ' Object representing user file

Dim logFSO         ' File system object to write log file
Dim logfileName    ' Log file name
Dim objLogFile     ' Object representing log file

Dim domain         ' The current domain
Dim dicUsers       ' A dictionary used to check for duplicate names
                   ' and store user's distinguished names

Dim preProcessedOK ' Flag to indicate that user names were all OK
Dim doIt           ' Flag to indicate that we should add users


'==========================================================
' Main program
'==========================================================
initialise
getCommandLineArguments
groupDN = getGroupDN( group, Domain )
If Not verifyFileExists( userFileName ) Then
  WScript.Echo "Cannot find user file"
  WScript.Quit
End If
checkCanCreateLogFile
' preProcess
openUserFile
  checkUserNames
closeUserFile
If Not preProcessedOK Then
  WScript.Echo "Processing failed - please fix errors in user file then try again"
  WScript.Echo "(No log file written)."
  WScript.Quit
End If
If doIt Then
  ' Make the additions
  openUserFile
    createLogFile
      addUsersToGroup
    closeLogFile
  closeUserFile
Else
  WScript.Echo "(No log file written)."
  WScript.Echo "Use /di switch to actually add users to group."
End If
WScript.Quit



'==========================================================
' Try to open users file
'==========================================================
Sub openUserFile
  On Error Resume Next
    Set userFSO = CreateObject( "scripting.FileSystemObject" )
    Set objUserFile = userFSO.OpenTextFile( userFileName )
    If Err.Number <> 0 Then
      On Error Goto 0
      WScript.Echo "Unable to open user file: " & userFileName
      WScript.Quit
    End If
End Sub



'==========================================================
' Try to create log file
'==========================================================
Sub createLogFile
  On Error Resume Next
    Set logFSO = CreateObject( "scripting.FileSystemObject" )
    Set objLogFile = logFSO.CreateTextFile( logFileName )
    If Err.Number <> 0 Then
      On Error Goto 0
      WScript.Echo "Unable to create log file for writing: " & logfileName
      WScript.Quit
    End If
End Sub



'==========================================================
' Close the user file
'==========================================================
Sub closeUserFile
  objUserFile.Close
End Sub



'==========================================================
' Close the log file
'==========================================================
Sub closeLogFile
  objLogFile.Close
End Sub



'==========================================================
' Check users in file exist and no duplicates
'==========================================================
Sub checkUserNames
  Dim userCount   ' Counter
  Dim userName    ' SAM name of current user
  Dim userDN      ' Distinguished name of current user
  Dim errorCount  ' Error count

  WScript.Echo "Validating names in user file."
  userCount = 0
  errorCount = 0
  Do Until objUserFile.AtEndOfStream
    userCount = userCount + 1
    userName = Trim( objUserFile.ReadLine )
    ' See if the name exists in AD
    If Not getUserDetails( userName, userDN ) Then
      ' User does not exist
      WScript.Echo "Line " & userCount & " - Non existent user : " & userName
      errorCount = errorCount + 1
    Else
      ' User name exists
      ' Have we seen him before?
      If dicUsers.Exists ( userName) Then
        ' Yes we have
        WScript.Echo "Line " & userCount & " - Duplicate user    : " & userName
        errorCount = errorCount + 1
      Else
        ' No we haven't - so add him to the dictionary.
        dicUsers.Add userName, userDN
      End If
    End If
  Loop
  WScript.Echo "  " & userCount & " lines processed. " & errorCount & " error(s)."
  If errorCount = 0 Then
    preProcessedOK = True
  Else
    preProcessedOK = False
  End If
End Sub


'==========================================================
' Add users to the group
'==========================================================
Sub addUsersToGroup
  Dim lineCount   ' Counts lines processed
  Dim addCount    ' Counts additions
  Dim userName    ' SAM name of current user
  Dim line        ' String to be written
  Dim objGroup    ' Object representing group

  ' Create group object
  Set objGroup = GetObject("LDAP://" & groupDN)

  line = "Adding users to group: " & group
  objLogFile.WriteLine( line )
  WScript.Echo line

  lineCount = 0
  addCount = 0
  Do Until objUserFile.AtEndOfStream
    userName = Trim( objUserFile.ReadLine )
    On Error Resume Next
    objGroup.Add( "LDAP://" & dicUsers( userName ) )
    If Err.Number <> 0 Then
      Err.clear
      On Error Goto 0
      line = "    Failed to add " & userName & " to " & group & "  <<<<<<<"
    Else
      line = "  Adding " & userName & " to " & group
      addCount = addCount + 1
    End If
    objLogFile.WriteLine( line )
    Wscript.Echo line
    lineCount = lineCount + 1
  Loop
  line = lineCount & " users processed " & addCount & " added to group."
  objLogFile.WriteLine( line )
  WScript.Echo line
End Sub


'==========================================================
' Given a SAM username the first parameter, it attempts
' to return the corresponding distinguished name in the
' second parameter.  The function returns true or false
' depending upon whether or not the user was found.
'==========================================================
Function getUserDetails( ByVal SAM, ByRef userDN )
  Const STR_SCOPE = "subtree"
  Const strAttributeList = "sAMAccountname,distinguishedName"

  Dim intRecordCount   ' Number of matching records found
  Dim objCommand       ' Used in building AD query
  Dim objConnection    ' Used in building AD query
  Dim strFilterText    ' Defines what to search for
  Dim strQueryText     ' Holds text of LDAP query
  Dim objRecordSet     ' Set of records returned by name look up

  ' Build query
  strFilterText = "(&(objectCategory=person)(objectClass=user)(sAMAccountname=" & SAM & "))"
  strQueryText = "<LDAP://" & domain & ">;" & strFilterText & ";" & strAttributeList & ";" & STR_SCOPE
  Set objCommand = CreateObject("ADODB.Command")
  Set objConnection = CreateObject("ADODB.Connection")
  objConnection.Provider = "ADsDSOObject"
  objConnection.Open "Active Directory Provider"
  objCommand.ActiveConnection = objConnection
  objCommand.CommandText = strQueryText
  ' WScript.Echo "Query text = " & strQueryText

  ' Execute query
  On Error Resume Next
  Set objRecordSet = objCommand.Execute
  If Err.Number <> 0 Then
    WScript.Echo vbTab & "LDAP query failed."
    WScript.Echo "Error number      = " & Err.Num
    WScript.Echo "Error description = " & Err.Description
    Err.clear
    WScript.Quit
  Else
    On Error Goto 0
    intRecordCount = 0
    Do Until objRecordSet.EOF
      intRecordCount = intRecordCount + 1
      userDN = objRecordSet.Fields("distinguishedName")
      objRecordSet.MoveNext
    Loop
  End If
  Select Case intRecordCount
    Case 0
      getUserDetails = False
    Case 1
      getUserDetails = True
    Case Else
      WScript.Echo "Internal error - More than one user found for a single name."
  End Select
End Function


'==========================================================
' Reads command line arguments
'==========================================================
Sub getCommandLineArguments
  Dim argCount
  Dim param
  Dim i

  argCount = WScript.Arguments.Count
  Select Case argCount
    Case 1, 2
      param = LCase( WScript.Arguments(0) )
      If param = "/ver" Then
        WScript.Echo "Program name: " & programName
        WScript.Echo "Version:      " & verNum
        WScript.Quit
      Else
        DisplayUsageInstructions
      End If
    Case 3, 4
      group        = LCase( WScript.Arguments(0) )
      userFileName = LCase( WScript.Arguments(1) )
      logfileName  = LCase( WScript.Arguments(2) )
      If argCount = 4 Then
        If LCase( WScript.Arguments( 3 ) ) = "/di" Then
          doIt = True
        End If
      End If
    Case Else
      DisplayUsageInstructions
  End Select
End Sub


'==========================================================
' Initialise globals
'==========================================================
Sub initialise
  ' Initialise user list - used to look for duplicates
  Set dicUsers = CreateObject("scripting.dictionary")
  ' Set domain
  domain = getDomain()
  preProcessedOK = False
  doIt           = False
End Sub



'==========================================================
' Establish current domain
'==========================================================
Function getDomain
  Dim objRootDSE
  Set objRootDSE = GetObject("LDAP://rootDSE")
  getDomain = objRootDSE.Get("defaultNamingContext")
End Function



'======================================================================================
' Write Usage instructions to the screen.
'======================================================================================
Sub DisplayUsageInstructions
  WScript.Echo "Program: autg.vbs (Add users to group)"
  WScript.Echo
  WScript.Echo "Purpose: Reads a list of usernames from a text file and adds them to the"
  WScript.Echo "         specified group in Active Directory."
  WScript.Echo
  WScript.Echo "         By default it validates the group name and user names, but does"
  WScript.Echo "         not actually add them.  To perform the add you must use the"
  WScript.Echo "         /di (DoIt) switch."
  WScript.Echo
  WScript.Echo "         The use of a log file is mandatory - to document what was done"
  WScript.Echo "         in case of the need to roll-back."
  WScript.Echo
  WScript.Echo "Usage:   autg  group_name  user_file_path  log_file_path  [ /di]"
  WScript.Echo
  WScript.Echo "Note:    This script will only work if the user running the script has the"
  WScript.Echo "         necessary rights to add the users to the group.  (No message is"
  WScript.Echo "         given if you don't - the additions simple fail.)"
  WScript.Echo
  WScript.Quit
End Sub



'==========================================================
' Given the group SAM name and the domain, searchs AD to
' and returned the distinguished name of the group.
' Terminates if cannot find the group.
'==========================================================
Function getGroupDN( ByVal groupShortName, ByVal strDomain )
  Dim intRecordCount   ' Number of matching records found
  Dim objCommand       ' Used in building AD query
  Dim objConnection    ' Used in building AD query
  Dim strFilterText    ' Defines what to search for
  Dim strQueryText     ' Holds text of LDAP query
  Dim objRootDSE       ' Used to get current domain
  Dim objRecordSet     ' Record set returned by LDAP query
  Dim tempFQDN         ' Temporarily holds FQDN
  Dim tempType         ' Temporarily holds group type

  Const strAttributeList = "distinguishedName,groupType,samAccountName"
  Const STR_SCOPE = "subtree"

  ' Build query
  strFilterText = "(&(objectCategory=group)(cn=" & groupShortName & "))"
  strQueryText = "<LDAP://" & strDomain & ">;" & strFilterText & ";" & strAttributeList & ";" & STR_SCOPE
  Set objCommand = CreateObject("ADODB.Command")
  Set objConnection = CreateObject("ADODB.Connection")
  objConnection.Provider = "ADsDSOObject"
  objConnection.Open "Active Directory Provider"
  objCommand.ActiveConnection = objConnection
  'Wscript.Echo "Query = " & strQueryText
  objCommand.CommandText = strQueryText

  ' Execute query
  On Error Resume Next
  Set objRecordSet = objCommand.Execute
  On Error Goto 0

  If Err.Number <> 0 Then
    ' Query failed
    Wscript.Echo "Internal error: Group name lookup failed."
    WScript.Echo "Error number = " & Err.Number
    WScript.Quit
  Else
    ' Query succeeded
    ' Count number of records returned
    Select Case objRecordSet.RecordCount
      Case 0
        WScript.Echo "Cannot find group: "& groupShortName
        WScript.Echo "Script terminating"
        WScript.Quit
      Case 1
        getGroupDN = objRecordSet.Fields( "distinguishedName" )
      Case Else
        WScript.Echo "Group name is ambiguous.  Please refine."
        intRecordCount = 0
        Do Until objRecordSet.EOF
          intRecordCount = intRecordCount + 1
          WScript.Echo  "   " & objRecordSet.Fields("distinguishedName")
         objRecordSet.MoveNext
        Loop
        WScript.Echo "Script terminating."
    End Select
  End If
End Function



'======================================================================================
' Check a file exists.
'======================================================================================
Function verifyFileExists( ByVal strInputFile )
  Dim objFSO
  Set objFSO = CreateObject( "scripting.FileSystemObject" )
  verifyFileExists = objFSO.FileExists( strInputFile )
End Function




'======================================================================================
' Check can create output file.
'======================================================================================
Sub checkCanCreateLogFile
  Dim errNum

  ' Check if file already exists
  If verifyFileExists( logfileName ) Then
    WScript.Echo "Specified output file already exists."
    WScript.Echo "Please specify the name of a new file."
    WScript.Quit
  End If
  ' If here, then file does not exist,
  ' so try to temporarily create it
  Set logFSO = CreateObject("Scripting.FileSystemObject")
  On Error Resume Next
    Set objLogFile = logFSO.CreateTextFile( logFileName )
    errNum = Err.Number
    Err.Clear
  On Error Goto 0
  If ErrNum <> 0 Then
    ' Couldn't make the file
    WScript.Echo "Unable to create log file: " & logFileName
    WScript.Quit
  Else
    ' Succeeded, so delete file.
    objLogFile.Close
    logFSO.DeleteFile( logfileName )
  End If
End Sub  