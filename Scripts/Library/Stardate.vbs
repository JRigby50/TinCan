' Time in Star Trek Stardate format
Option Explicit		'force explicit variable declaration
On Error Resume Next	'on error continue at next statement

' Initialize variables
Dim timerID, now, nowMonth, nowDate, nowYear, nowHour, nowMinute, nowSecond, Elapsed, WshShell, CRLF
Dim thisYear, lpyra, lpyrb, issue, yearsPast, total, mns, scs, temp2, doLoop, btnCode, insStr
timerID = 0
CRLF =  Chr(13) & Chr(10)
doLoop = TRUE
Set WshShell = CreateObject( "WScript.Shell" )
insStr = CRLF & CRLF & "Press OK to Stop"

' Main
While doLoop = TRUE
  If stardate() = 1 Then  ' drop out of loop if OK clicked
     doLoop = FALSE
  End If
WEnd

Function stardate()
  now = Date()
  nowMonth = Month(now) - 1 'subtract cos so Jan=0 not 1 to be consistent with getMonth Vb function
  nowDate = Day(now)
  nowYear = Year(now) + 1900
  now = Time()
  nowHour = Hour(now)
  nowMinute = Minute(now)
  nowSecond = Second(now)
  now = 0
  Elapsed = nowSecond + 60 * (nowMinute) + 3600 * (nowHour) +86400 * (nowDate - 1)
  If (nowMonth>10) Then
        Elapsed = Elapsed + (86400*334)
   Else If (nowMonth>9) Then 
        Elapsed = Elapsed + (86400*304)
    Else If (nowMonth>8) Then
           Elapsed = Elapsed + (86400*273)
      Else If (nowMonth>7) Then
            Elapsed = Elapsed + (86400*243)
        Else If (nowMonth>6) Then 
               Elapsed = Elapsed + (86400*212)
          Else If (nowMonth>5) Then
                 Elapsed = Elapsed + (86400*181)
            Else If (nowMonth>4) Then
                   Elapsed = Elapsed + (86400*151)
              Else If (nowMonth>3) Then
                     Elapsed = Elapsed + (86400*120)
                Else If (nowMonth>2) Then
                     Elapsed = Elapsed + (86400*90)
                  Else If (nowMonth>1) Then
                         Elapsed = Elapsed + (86400*59)
                    Else If (nowMonth>0) Then
                             Elapsed = Elapsed + (86400*31)
                         End If
                       End If
                     End If
                   End If
                 End If
               End If
             End If
           End If
         End If
        End If
  End If
  If (nowYear>2100) Then
       nowYear = nowYear-1900
  End If
  thisYear = Round( Elapsed /  315.36) / 100
  lpyra= Round(nowYear/400)
  lpyrb= nowYear/400

  If (lpyra=lpyrb) Then
    If (nowMonth>2) Then
      Elapsed = Elapsed + (86400)
    End If
  End If

  issue = Round(((nowYear-2323)/100)-.5)
  yearsPast = (nowYear - (2323+(issue * 100))) * 1000
  total = thisYear+yearsPast
  If (nowMinute<10) Then 
      mns="0"
    Else mns=""
  End If
  If (nowSecond<10) Then 
      scs="0"
    Else scs=""
  End If
  temp2 = "[" & issue & "]  " & total & "    " & nowHour & ":" & mns & nowMinute & ":" & scs & nowSecond & insStr
  stardate = WshShell.Popup( temp2, 1, "StarDate", 0)
End Function


