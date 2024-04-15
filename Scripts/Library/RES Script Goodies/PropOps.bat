@Echo off
CLS
Echo This is the RESVA Site PropOps Logon Script...
REM - Last Modified: 08/29/2002 - SMG

REM - Setting Drive & Share Variables -
SET Drive = O:
SET Share1 = \\RESVA-PRO1\AUA_TAC

REM - Create Drive Mappings
NET use %Drive /delete && CLS
NET use %Drive %Share1%
IF Errorlevel = 0 CLS && ECHO "MAP Success" && Goto END

CLS && ECHO "ERROR Mapping Drive %Drive% to %Share1%

:END