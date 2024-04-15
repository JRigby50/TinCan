'Generate a unique file name for each backup based on the current date.

ThisDay = Day(Date)
StrThisDay = ""
If ThisDay < 10 Then
StrThisDay = "0" & ThisDay
Else StrThisDay = ThisDay
END If

ThisMonth = Month(Date)
StrThisMonth = ""
If ThisMonth < 10 Then
StrThisMonth = "0" & ThisMonth
Else StrThisMonth = ThisMonth
END If

ThisYear = Year(Date)

StrBackupName = "AccessControl_db_" & ThisYear & StrThisMonth & StrThisDay & "0200.BAK"

MsgBox strBackupName