Sub Ringin
  Set objShell = CreateObject("Wscript.Shell")
  Set WshSysEnv = objShell.Environment("PROCESS")  
  strSoundFile = WshSysEnv("WINDIR") & "\Media\ringin.wav" 
  strCommand = "sndrec32 /play /close " & chr(34) & strSoundFile & chr(34)
  objShell.Run strCommand, 0, False    
  Wscript.Sleep 100
  Set objShell = Nothing
End Sub

Dim WshShell, BtnCode

Set WshShell = WScript.CreateObject("WScript.Shell")

Do Until BtnCode = 7
  Call Ringin
  BtnCode = WshShell.Popup("Snooze for 1 minute?", 5, "Reminder:", 4 + 32)
  If BtnCode = 6 Then 
    Wscript.Sleep 60000
  End If
Loop
