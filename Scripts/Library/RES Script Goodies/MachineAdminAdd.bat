@ECHO OFF
REM - MachineAdminAdd.bat - Add the Domain Admins and Local Machine Admins groups
REM - to target machine's local administrator group.
REM - Created by: Michael Greelish (s823755) - RESVA
REM - Last Modified: 10-10-2003

CLS
:GROUP_WRK1
REM Checking for INFO\Domain Admin Group in local Administrator Group - Adding if missing!
net localgroup Administrators | find /i "%USERDOMAIN%\DOMAIN ADMINS" > nul

REM If not, the group is added
if not %errorlevel%==0 goto :ADDGROUP1
goto :GROUP_WRK2

CLS
:ADDGROUP1
net localgroup Administrators "%USERDOMAIN%\DOMAIN ADMINS" /ADD > nul

CLS
:GROUP_WRK2
REM Checking for INFO\Local Machine Admin Group in local Administrator Group - Adding if missing!
net localgroup Administrators | find /i "%USERDOMAIN%\INFO GG MACHINE ADMINS" > nul

REM If not, the group is added
if not %errorlevel%==0 goto :ADDGROUP2
goto :END

CLS
:ADDGROUP2
net localgroup Administrators "%USERDOMAIN%\INFO GG MACHINE ADMINS" /ADD > nul

CLS
:END