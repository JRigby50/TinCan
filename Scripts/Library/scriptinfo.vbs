WScript.Echo GetSEInfo()
WScript.Echo GetScript()
Function GetSEInfo
Dim info
info = ""
info = ScriptEngine & " Version "
info = info & ScriptEngineMajorVersion & "."
info = info & ScriptEngineMinorVersion & "."
info = info & ScriptEngineBuildVersion
GetSEInfo = info
End Function
Function GetScript
Dim info
scr = "Name: "
scr = WScript.ScriptName & " Full path: "
scr = scr & WScript.ScriptFullName
GetScript = scr
End Function
