@echo off
' ======================================================================
'
' NAME: DHCPBackup_scopes.bat
'
' AUTHOR: Michael J. Ginter
' DATE  : 1/29/2008
' 
'	Version: 1.0
' 
' COMMENT: Creates backups of each DHCP Scope individually.  Will auto-
'	create a CABS folder in the directory run to store output file.
'	**** The DHCP Service will stop and restart during backup.
' 
' **DISCLAIMER**
'    THIS MATERIAL IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
'    EITHER EXPRESS OR IMPLIED, INCLUDING, BUT Not LIMITED TO, THE
'    IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'    PURPOSE, OR NON-INFRINGEMENT. SOME JURISDICTIONS DO NOT ALLOW THE
'    EXCLUSION OF IMPLIED WARRANTIES, SO THE ABOVE EXCLUSION MAY NOT
'    APPLY TO YOU. IN NO EVENT WILL I BE LIABLE TO ANY PARTY FOR ANY
'    DIRECT, INDIRECT, SPECIAL OR OTHER CONSEQUENTIAL DAMAGES FOR ANY
'    USE OF THIS MATERIAL INCLUDING, WITHOUT LIMITATION, ANY LOST
'    PROFITS, BUSINESS INTERRUPTION, LOSS OF PROGRAMS OR OTHER DATA ON
'    YOUR INFORMATION HANDLING SYSTEM OR OTHERWISE, EVEN If WE ARE
'    EXPRESSLY ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. 
'
' ======================================================================
' Create Date Variable (US Only) YYYYMMDD
:DATE
For /F "Tokens=2" %%I in ('Date /t') Do Set DTemp=%%I
For /F "Delims=/,= Tokens=2" %%I in ('Set DTemp') Do Set TDate=%%I
For /F "Delims=/,= Tokens=3" %%I in ('Set DTemp') Do Set TDate=%TDate%%%I
For /F "Delims=/,= Tokens=4" %%I in ('Set DTemp') Do Set TDate=%%I%TDate%
Set DTemp=

' Current Directory
set DHCPBACKUPPATH=%~dp0

' Directory of where to store DHCP Backup Cab files
set CABDIR=%DHCPBACKUPPATH%cabs

' Create Directory to store cab files if it does not exist
if not exist %CABDIR% md %CABDIR%

' Create a DATED directory to store today's DHCP Backup
if not exist %DHCPBACKUPPATH%%TDATE% md %DHCPBACKUPPATH%%TDATE%

' CAB directive file (.DDF) needed to cab all files backed up in to one Cab file
Set CABDIRFILE=%DHCPBACKUPPATH%%TDATE%\DHCPBackup.DDF

' Backup each dhcp scope individually on local system to dated folder
For /f "TOKENS=1 SKIP=4" %%I In ('netsh dhcp server \\%COMPUTERNAME% show scope ^|find /i "-"') do netsh dhcp server 
\\%COMPUTERNAME% export %DHCPBACKUPPATH%%TDATE%\dhcpdb_%%I_%COMPUTERNAME%.dhcp %%I >NUL

' Create a dump of the current DHCP server settings to a txt file
netsh dhcp server \\%COMPUTERNAME% dump > %DHCPBACKUPPATH%%TDATE%\dhcp_%COMPUTERNAME%_config.txt.dhcp

' Create DDF (cab directive file) heading
echo ;***%COMPUTERNAME% DHCP Backup %TDATE% Directive file >%CABDIRFILE%
echo ; >>%CABDIRFILE%
echo .OPTION EXPLICIT >>%CABDIRFILE%
echo .Set CabinetNameTemplate=%COMPUTERNAME%_DHCPBackup_%TDATE%.CAB >>%CABDIRFILE%
echo .set DiskDirectoryTemplate=%CABDIR% >>%CABDIRFILE%
echo .Set MaxDiskSize=CDROM >>%CABDIRFILE%
echo .Set FolderSizeThreshold=2000000 >>%CABDIRFILE%
echo .Set CompressionType=MSZIP >>%CABDIRFILE%
echo .Set Cabinet=on >>%CABDIRFILE%
echo .Set Compress=on >>%CABDIRFILE%

' add files to be cab'ed to the DDF
for /f %%I in ('dir %DHCPBACKUPPATH%%TDATE%\*.dhcp /b') do @echo %DHCPBACKUPPATH%%TDATE%\%%I >>%CABDIRFILE%

' switch to the dated folder and create cab file
cd /d %DHCPBACKUPPATH%%TDATE% & makecab /f %CABDIRFILE% /L %CABDIR% >NUL

' switch to script root directory and remove the dated folder and all files
cd /d %DHCPBACKUPPATH% & rd /s /q %DHCPBACKUPPATH%%TDATE% >NUL

