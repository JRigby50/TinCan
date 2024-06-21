@echo off

set ArchiveDirectory=C:\Documents and Settings\s823753\Desktop\Daily Logs\EventLogs
echo .
echo .
rem --------------------------------------------------------------------------------


kix32 GetEventLogs.kix $ArchiveDirectory="%ArchiveDirectory%"
