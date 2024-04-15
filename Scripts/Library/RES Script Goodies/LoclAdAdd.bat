@ECHO OFF

REM Created by: Michael Greelish (s823755) - RESVA
REM Last Modified: 06-10-2003

CLS

:GO_SYSTEMROOT
chdir /D %SystemRoot%\system32

:LOC_ADMN_SET
REM Checking for Local User in Local Administrators Group - Adding if missing!
net localgroup Administrators | find /i "Administrator" > nul


REM If not, 
if not %errorlevel%==0 goto :GROUP_WRK1

REM Reset Local Administrator account password
net USER Administrator Chantilly1 > nul

CLS
:GROUP_WRK1
REM Checking for Local OU Admin Group - Adding if missing!
net localgroup Administrators | find /i "%USERDOMAIN%\RESVA GG SITE ADMINS" > nul

REM If not, the group is added
if not %errorlevel%==0 goto :ADDGROUP1
goto :GROUP_WRK2

CLS
:ADDGROUP1
net localgroup Administrators "%USERDOMAIN%\RESVA GG SITE ADMINS" /ADD > nul

CLS
:GROUP_WRK2
REM Checking for Local OU Admin Group - Adding if missing!
net localgroup Administrators | find /i "%USERDOMAIN%\RESVA GG SITE TECHS" > nul

REM If not, the group is added
if not %errorlevel%==0 goto :ADDGROUP2
goto :LOC_USER_WRK

CLS
:ADDGROUP2
net localgroup Administrators "%USERDOMAIN%\RESVA GG SITE TECHS" /ADD > nul

CLS
:LOC_USER_WRK
REM Checking for Local User in Local Administrators Group - Adding if missing!
net localgroup Administrators | find /i "%USERNAME%" > nul

REM If not, the user is added
if not %errorlevel%==0 goto :ADDUSER1
goto :DOM_USER_WRK

CLS
:ADDUSER1
net localgroup Administrators "%USERNAME%" /ADD > nul

CLS
:DOM_USER_WRK
REM Checking for Domain User in Local Administrators Group - Adding if missing!
net localgroup Administrators | find /i "%USERDOMAIN%\%USERNAME%" > nul

REM If not, the user is added
if not %errorlevel%==0 goto :ADDUSER2
goto :END

CLS
:ADDUSER2
net localgroup Administrators "%USERDOMAIN%\%USERNAME%" /ADD > nul

CLS
:END
CLS
