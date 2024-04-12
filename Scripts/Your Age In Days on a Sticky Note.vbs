'Your Age In Days on a Sticky Note

firstname="Jim"
birthdate=DateSerial(1951,7,4)

days=DateDiff("d",birthdate,Date)

Set img=CreateObject("vbsedit.imageprocessor")

Set objShell = CreateObject( "WScript.Shell" )
resourceLocation=objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Adersoft\Vbsedit\Resources\"

img.Load resourceLocation & "calendar.png"
img.FontFamily = "Arial"
img.FontSize = 12
img.Color="White"

img.CenterText firstname,0,15,130,22

img.FontSize = 26
img.CenterText days,0,37,130,60

img.FontFamily = "Arial"
img.FontSize = 9
img.CenterText "your age in days",0,90,130,20
         
path = resourceLocation & "\ageindays_" &  firstname & ".png"
img.Save path

Set note = img.CreateStickyNote("ageindays_" & firstname,path,"#00FF00")
  
note.AddMenuOption "Edit Script with Vbsedit","""c:\program files\vbsedit\vbsedit.exe"" """ & WScript.ScriptFullName & """"
note.AddMenuOption "Edit Script with Notepad","notepad.exe """ & WScript.ScriptFullName & """"
note.AddMenuOption "Refresh","wscript.exe //B """ & WScript.ScriptFullName & """"
  
note.ShowBalloon firstname & " is " & days & " days old","My age in days",0

