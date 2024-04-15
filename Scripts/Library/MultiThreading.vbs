Option Explicit
'* -----------------------------------------------------------------------------
'* Define Constants
'*
Const ForReading				= 1
Const ForAppending				= 8 
Const adVarChar					= 200
Const adInteger					= 3
Const adFilterNone				= ""
Const MaxCharacters				= 255
'* change MaxThreads to the number of cscripts you want to be running simultaneously
'* NOTE: there can actually be a MaxThreads + 2 script engines running
Const MaxThreads				= 15

'* -----------------------------------------------------------------------------
'* Declare variables
'*
dim g_ador
dim g_fso
dim g_ts
dim infile
dim oWshShell
dim uargs 
dim sTarget 
dim cThread 
dim aExec()
dim logFile
dim i 

'* the output 	
logFile = "C:\Output\myLog.log"

'* use argument as the file path for input
set uargs = wscript.arguments.unnamed

if uargs.count = 0 then 
	wscript.quit(0) 
end if

set g_fso = createobject("scripting.filesystemobject")
set infile = g_fso.OpenTextFile(trim(uargs(0)), ForReading )
Set g_ts = g_fso.OpenTextFile(logFile, ForAppending, true)

g_ts.writeline Now & " | script called Target: " & uargs(0) 

'* create Disconnected RS
Set g_ador = CreateObject("ADOR.Recordset")
g_ador.Fields.Append "ComputerName", adVarChar, MaxCharacters
g_ador.Fields.Append "ExecStatus", adInteger, 5 
g_ador.Fields.Append "hExec", adInteger, 5
g_ador.Open

Set oWshShell = WScript.CreateObject("WScript.Shell")

'* count of script engines executing
cThread = 0 
i = 0 

do until infile.atendofstream
	sTarget = trim(infile.ReadLine)
	if sTarget <> "" then
		
		redim preserve aExec(i) 
		'* script.vbs is a script that returns and exit code using wscript.quit(1) or wscript.quit(0)
		'* the value of the exit code will be used as the results for logging
		'* the call to the script will be script.vbs arg
		'* this script could do anything such as return a ping result or check a machine for a certian
		'* software.  I suppose you could even build a sort of enum and depending on your exit codes
		'* return certian strings from the main script in a select case statment
		set aExec(i) = oWshShell.Exec("wscript.exe ""C:\scripts\script.vbs"" " & sTarget)
		
		g_ador.addnew
		g_ador("ComputerName") = sTarget
		g_ador("ExecStatus") = 0 
		g_ador("hExec") = i
		g_ador.Update 
		cThread = cThread + 1 
		if cThread > MaxThreads then 
			WaitForThreads(MaxThreads)
			g_ador.MoveLast
		end if
		i = i + 1 
	end if
loop

infile.close 

'* wait for the remaining threads to quit
WaitForThreads(0)

g_ts.close
g_ador.Close 

'* -----------------------------------------------------------------------------
'* Sub: WaitForThreads
'*
'* Purpose:  Wait for number of executions to drop under MAX before
'*			 executing another thread (script engine)
'*
'* Input:    [in] Max  the number of threads running
'*
'* Output:   none 
'*
'* -----------------------------------------------------------------------------
Public Sub WaitForThreads(byVal Max)

	dim n		  '* the handle number of the still running oWshShell.Exec

	'* once the cThread drops below max the main loop will continue and start 
	'* executing more script engines 
	do while cThread > Max
		
		'* give scripts a moment to return results
		wscript.sleep 200
		'* filter ador for records that we do not have results for yet
		'* when a result is recieved the ExecStatus is changed to 1
		g_ador.Filter = "ExecStatus = 0" 
		g_ador.MoveFirst 
		while not g_ador.EOF 
			n = g_ador("hExec") 
			if aExec(n).Status <> 0 then 
				'* update status to complete
				g_ador("ExecStatus") =  aExec(n).Status
				g_ts.writeline Now & " | " & g_ador("ComputerName") & " | Return: " &  aExec(n).ExitCode
			end if
			g_ador.MoveNext 
		wend 
		cThread = g_ador.RecordCount 
	loop
	
	'* remove filter
	g_ador.Filter = adFilterNone
	
end sub
