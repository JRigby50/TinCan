@echo off
REM Da ha
REM 10/5/2005
REM Change proxy settings

SET PROXY1=http://autoproxy1.northgrum.com/proxy/bceast.pac
SET PROXY2=http://autoproxy2.northgrum.com/proxy/bccentral.pac
SET PROXYCFG=http://proxycfg.it.northgrum.com/proxy/ngit.pac
SET MARYLAND=http://www.md.essd.northgrum.com/proxy/essd.pac

c:\proxy\regfind -p "HKEY_CURRENT_USER\Software\Microsoft\Windows\Currentversion\Internet Settings" -n AutoConfigURL > C:\test.txt
cls
echo.
echo ...Your current proxy setting is
FOR /F "tokens=2* delims== " %%i IN (C:\test.txt) DO (
	IF %%i==%PROXY1% echo %%i%
	if %%i==%PROXY2% echo %%i%
	if %%i==%PROXYCFG% echo %%i%
	if %%i==%MARYLAND% echo %%i%)
echo.


echo To change your proxy setting, choose one of the following 5 choices.
echo.
echo A. http://autoproxy1.northgrum.com/proxy/bceast.pac
echo B. http://autoproxy2.northgrum.com/proxy/bccentral.pac
echo C. http://proxycfg.it.northgrum.com/proxy/ngit.pac
echo D. http://www.md.essd.northgrum.com/proxy/essd.pac
echo E. Exit out. I don't want to make any changes at this time.
echo.

c:\proxy\CHOICE /c:abcde
IF ERRORLEVEL 5 GOTO END
IF ERRORLEVEL 4 GOTO MARYLAND
IF ERRORLEVEL 3 GOTO PROXYCFG
IF ERRORLEVEL 2 GOTO PROXY2
IF ERRORLEVEL 1 GOTO PROXY1

:PROXY1
REG ADD "HKCU\Software\Microsoft\Windows\Currentversion\Internet Settings" /v AutoConfigURL /t REG_SZ /d http://autoproxy1.northgrum.com/proxy/bceast.pac /f
echo.
echo Please close any current open Internet Explorer Windows and re-open it.
echo.
PAUSE
GOTO END

:PROXY2
REG ADD "HKCU\Software\Microsoft\Windows\Currentversion\Internet Settings" /v AutoConfigURL /t REG_SZ /d http://autoproxy2.northgrum.com/proxy/bccentral.pac /f
echo.
echo Please close any current open Internet Explorer Windows and re-open it.
echo.
PAUSE
GOTO END

:PROXYCFG
REG ADD "HKCU\Software\Microsoft\Windows\Currentversion\Internet Settings" /v AutoConfigURL /t REG_SZ /d http://proxycfg.it.northgrum.com/proxy/ngit.pac /f
echo.
echo Please close any current open Internet Explorer Windows and re-open it.
echo.
GOTO END

:MARYLAND
REG ADD "HKCU\Software\Microsoft\Windows\Currentversion\Internet Settings" /v AutoConfigURL /t REG_SZ /d http://www.md.essd.northgrum.com/proxy/essd.pac /f
echo.
echo Please close any current open Internet Explorer Windows and re-open it.
echo.
PAUSE
GOTO END

:END
EXIT
