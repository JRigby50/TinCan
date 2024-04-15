if exist \\rslv26-iismgmt2\dr\DHCP\Exports\old\%computername%_dhcp_export.txt goto delold
if not exist \\rslv26-iismgmt2\dr\DHCP\Exports\old\%computername%_dhcp_export.txt goto checkold

:delold
del \\rslv26-iismgmt2\dr\DHCP\Exports\old\%computername%_dhcp_export.txt
goto moveold

:checkold
if exist \\rslv26-iismgmt2\dr\DHCP\Exports\%computername%_dhcp_export.txt goto moveold
goto export

:moveold 
copy \\rslv26-iismgmt2\dr\DHCP\Exports\%computername%_dhcp_export.txt \\rslv26-iismgmt2\dr\DHCP\Exports\old
del \\rslv26-iismgmt2\dr\DHCP\Exports\%computername%_dhcp_export.txt
goto export

:export
netsh dhcp server export \\rslv26-iismgmt2\dr\DHCP\Exports\%computername%_dhcp_export.txt all
goto end

:end