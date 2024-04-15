:SAVSTART
@ECHO OFF
REM RESVA Site Specific Symantec GRC.Dat Update Section
REM This will insure all RES OU users recieve the current GRC.dat SAV-Managed file
REM Note - VAR1/RES is using GRC.res as a watermark
CLS
if exist "C:\Documents and Settings\All Users\Application Data\Symantec\Norton AntiVirus Corporate Edition\7.5\GRC.res" goto  SAVEND
Copy \\resva-snw02\VPLOGON\grc.* "C:\Documents and Settings\All Users\Application Data\Symantec\Norton AntiVirus Corporate Edition\7.5"
Copy %logonserver%\netlogon\shutdown.exe "%systemroot%\system32"
CLS
:SAVEND