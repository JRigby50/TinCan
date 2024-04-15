@ECHO OFF

Echo This is the RESVA Site Logon Script...
REM - Last Modified: 07/29/2003 - SMG
REM - Remove Icollect pilot install
REM - Last Modified: 09/08/2003 - SMG
REM - Added Symantec GRC.dat update section


:ICSTART
REM ICollect Installation Section
call %logonserver%\netlogon\ic.bat
:ICEND

:MIGSTART
REM Outlook Migration Profile Section
%logonserver%\netlogon\ifmember "RESVA GG MIGRATION"
if not errorlevel 1 goto migend
if exist "%userprofile%\water.mrk" goto migend
Copy %logonserver%\netlogon\water.mrk "%userprofile%"
%logonserver%\netlogon\profile.exe
:MIGEND

REM -<>----------------------------------<>-
REM Add Site Specific Entries Starting Here!
REM -<>----------------------------------<>-

REM --- Start - Check or ADD Default Admin Accounts ---
call %logonserver%\netlogon\resva\AdminAdd.bat
REM --- End - Check or ADD Default Admin Accounts ---

:SMSSTART
REM System Management Server 2.0 Section
%logonserver%\netlogon\ifmember "RESVA GG SMSUSERS"
if not errorlevel 1 goto SKIPSMS
call %logonserver%\netlogon\smsls.bat
:SKIPSMS
:SMSEND

:PRNSTART
REM W2K migration team added 8/7/02
net use S: \\resva-fsv04\shares1 /PERSISTENT:NO
CSCRIPT %logonserver%\netlogon\resva\MKPM.vbs S:\ 
REM end W2K migration team modification
:PRNEND

:PROPSTART
REM Proposal Operations Section
%logonserver%\netlogon\ifmember "RESVA GG PropOps"
if not errorlevel 1 goto PROPEND
call %logonserver%\netlogon\resva\PropOps.bat
:PROPEND

:SAV-GRC-START
call %logonserver%\netlogon\resva\SAVinst.bat
:SAV-GRC-END

goto end
:END