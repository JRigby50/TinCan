Option Explicit
'===========================================================================
'  Scheduled Task - Visual Basic ActiveX Script
'===========================================================================
'  Date    : 09/12/2002
'  Auth    : Israel Farfan
'  Desc    : VBScript file that sends short-on-disk warnings
'===========================================================================

'--  Check that server has minimum disk space (in MBs)
Call CheckDrives(300)




'===========================================================================
'  Name    : CheckDrives(intMinMBSize)
'  Desc    : Check Disk Space on Fixed Drives and report on short space
'===========================================================================
Function CheckDrives(intMinMBSize)
   Dim oFileSystem, oDrive, oDriveCollection
   Dim blnShortSpace, strErrorMsg, intMBSize

   blnShortSpace = False
   strErrorMsg = ""

   '-- Check each Fixed Drive on this Machine
   Set oFileSystem = CreateObject("Scripting.FileSystemObject")
   Set oDriveCollection = oFileSystem.Drives
   For Each oDrive in oDriveCollection
      '-- Check Fixed Drives Only
      If oDrive.DriveType = 2 Then
         '-- Get Disk Size as MBs (Bytes / 1024 = KBytes, KBytes / 1024 = MBytes)
         intMBSize = oDrive.FreeSpace / 1048576

         '-- Rise a Flag when less than Minimum
         If intMBSize < intMinMBSize Then
            blnShortSpace = True
            strErrorMsg = strErrorMsg & "Drive " & oDrive.DriveLetter & 
                ": - Free Space Left is " & FormatNumber(intMBSize, 2) & " MB" & vbCrLf
         End If
      End If
   Next
   Set oDriveCollection = Nothing
   Set oFileSystem = Nothing

   '-- Send Email if there is short space left on this machine
   If blnShortSpace = True Then Call SendAlertEmail("Running Low on Disk Space", _
       strErrorMsg)
End Function


'===========================================================================
'  Name    : GetMachineName()
'  Desc    : Determine Computer's Machine Name
'===========================================================================
Function GetMachineName()
    On Error Resume Next

    '-- Try Network Object
    Dim WshNetwork

    Set WshNetwork = CreateObject("WScript.Network")
    GetMachineName = WshNetwork.ComputerName
    Set WshNetwork = Nothing

    If IsNull(GetMachineName) Or GetMachineName = "" Then
        '-- Try Shell Object
        Dim wshShell

        Set wshShell = CreateObject("WSCript.Shell")
        GetMachineName = wshShell.RegRead _
            ("HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName")
        Set wshShell = Nothing

        If IsNull(GetMachineName) Or GetMachineName = "" Then
            '-- Try Windows Mgmt Implementation
            Dim objWMIService, OSItems, objItem
        
            Set objWMIService = GetObject _
               ("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
            Set OSItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem", , 48)
            For Each objItem in OSItems
            GetMachineName = objItem.CSName
            Next
        End If
    End If
    Err.Clear

    If IsNull(GetMachineName) Or GetMachineName = "" Then GetMachineName = "LOCALHOST"
End Function


'===========================================================================
'  Name    : SendAlertEmail(Subject Line, Email Body)
'  Desc    : Send Error/Alert Email
'===========================================================================
Function SendAlertEmail(strSubject, strErrorText)
    '-- Get Machine Name
    Dim oEmail, strMachineName
    strMachineName = GetMachineName()

    '-- Send Email
    Set oEmail = CreateObject("CDO.Message")
    oEmail.From = strMachineName
    oEmail.To   = "kmyer@fabrikam.com"
    oEmail.Subject = strMachineName & " : " & strSubject & " @ " & Now
    oEmail.TextBody = strErrorText
    oEmail.Send()
    Set oEmail = Nothing
End Function

