Set objArgs = WScript.Arguments
For I = 0 to objArgs.Count - 1
  WScript.Echo "File " + CStr(I) + ": " + objArgs(I)
Next