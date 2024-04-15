Option Explicit
On Error Resume Next
Dim strMoniker
Dim refWMI
Dim colEventLogs
Dim refEventLog
Dim strSource

'moniker string stub - security privilege needed to get
'numrecords for Security log
strMoniker = "winMgmts:{(Security)}!"

'append to moniker string if a machine name has been given
If WScript.Arguments.Count = 1 Then _
strMoniker = strMoniker & "\\" & WScript.Arguments(0) & ":"

'attempt to connect to WMI
Set refWMI = GetObject(strMoniker)
If Err <> 0 Then
WScript.Echo "Could not connect to the WMI service."
WScript.Quit
End If

'get a collection of Win32_NTEventLogFile objects
Set colEventLogs = refWMI.InstancesOf("Win32_NTEventLogFile")
If Err <> 0 Then
WScript.Echo "Could not retrieve Event Log objects"
WScript.Quit
End If

'iterate through each log and output information
For Each refEventLog In colEventLogs
WScript.Echo "Information for the " & _
refEventLog.LogfileName & _
" log:"
WScript.Echo " Current file size: " & refEventLog.FileSize
WScript.Echo " Maximum file size: " & refEventLog.MaxFileSize
WScript.Echo " The Log currently contains " & _
refEventLog.NumberOfRecords & " records"

'output policy info in a friendly format using OverwriteOutDated,
'as OverWritePolicy is utterly pointless.
'note "-1" is the signed interpretation of 4294967295
Select Case refEventLog.OverwriteOutDated
Case 0 WScript.Echo _
" Log entries may be overwritten as required"
Case -1 WScript.Echo _
" Log entries may NEVER be overwritten"
Case Else WScript.Echo _
" Log entries may be overwritten after " & _
refEventLog.OverwriteOutDated & " days"
WScript.Echo
End Select
Next

Set refEventLog = Nothing
Set colEventLogs = Nothing
Set refWMI = Nothing