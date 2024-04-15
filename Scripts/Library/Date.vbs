Dim sVar
sVar = InputBox("Go ahead, type something.","Test","Something")
MsgBox "The first letter is " & Left(sVar, 1)
MsgBox "and the last letter is " & Right(sVar, 1)
MsgBox "In uppercase it’s " & UCase(sVar)
MsgBox "Today is " & Date()
MsgBox "Right now it is " & Now()
MsgBox "The second character is " & Right(Left(sVar,2),1)