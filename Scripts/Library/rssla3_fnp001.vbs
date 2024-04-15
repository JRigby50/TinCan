
'Author - Randy Foster
'Email - rfoster_2000@excite.com
'MCSE NT4, 2000
'MCSA 2003
'MCDBA SQL7, SQL2000
'CCNA


'This script works well in a login script. When ran from a workstation
'it will loop thru all of the printers installed for the current logged on
'user and if they are connected to the old server they will be removed
'and then remaped to the new server. It will also attempt to descover
'which printer is the default printer in order to reset this if it is one
'of the printers that will be moved.
'
'Note: All of the printers must exist on the new server before this script
' should be run. If not then the printer will simply be removed from
' users profile and the script will not be able to reconnect them.
' "Print Migrator" is a utility that is part of the Windows 2000 and
' Windows 2003 resource kit. This will create all of the printers
' on the new server while maintaining all of their settings including
' print queue security. Keep in mind that for a time you will see
' duplicate printer if you do a search on printers, one advertisec
' from each server.
'
'We used this method to move just over 100 printers from one server to another
'and then ran this script as a logon script in a group policy. We let it run for
'a few days in order to allow for some users who don't logout every day have a
'chance to run it. All of our users profiles were updated and they didn't even
'know it was happening. We then deleted all of the printers from the old server.
'
'Note: The script does not run if you are Terminal serviced or SMS remote controling.



Option Explicit
Dim from_sv, to_sv, PrinterPath, PrinterName, DefaultPrinterName, DefaultPrinter
Dim DefaultPrinterServer, SetDefault, key
Dim spoint, Loop_Counter
Dim WshNet, WshShell
Dim WS_Printers
DefaultPrinterName = ""
spoint = 0
SetDefault = 0
set WshShell = CreateObject("WScript.shell")

from_sv = "\\pca4423" 'This should be the name of the old server.
to_sv = "\\rssla3-fnp001" 'This should be the name of your new server.

'Just incase their are no printers and therefor no defauld printer set
' this will prevent the script form erroring out.
On Error Resume Next
key = "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Device"
DefaultPrinter = LCase(WshShell.RegRead (key))
If Err.Number <> 0 Then
DefaultPrinterName = ""
else
'If the registry read was successful then parse out the printer name so we can
' compare it with each printer later and reset the correct default printer
' if one of them matches this one read from the registry.
spoint = instr(3,DefaultPrinter,"\")+1
DefaultPrinterServer = left(DefaultPrinter,spoint-2)
if DefaultPrinterServer = from_sv then
DefaultPrinterName = mid(DefaultPrinter,spoint,len(DefaultPrinter)-spoint+1)
end if
end if
Set WshNet = CreateObject("WScript.Network")
Set WS_Printers = WshNet.EnumPrinterConnections
'You have to step by 2 because only the even numbers will be the print queue's
' server and share name. The odd numbers are the printer names.
For Loop_Counter = 0 To WS_Printers.Count - 1 Step 2
'Remember the + 1 is to get the full path ie.. \\your_server\your_printer.
PrinterPath = lcase(WS_Printers(Loop_Counter + 1))
'We only want to work with the network printers that are mapped to the original
' server, so we check for "\\Your_server".
if LEFT(PrinterPath,len(from_sv)) = from_sv then
'Now we need to parse the PrinterPath to get rhe Printer Name.
spoint = instr(3,PrinterPath,"\")+1
PrinterName = mid(PrinterPath,spoint,len(PrinterPath)-spoint+1)
'Now remove the old printer connection.
WshNet.RemovePrinterConnection from_sv+"\"+PrinterName
'and then create the new connection.
WshNet.AddWindowsPrinterConnection to_sv+"\"+PrinterName
'If this printer matches the default printer that we got from the registry then
' set it to be the default printer.
if DefaultPrinterName = PrinterName then
WshNet.SetDefaultPrinter to_sv+"\"+PrinterName
end if
end if
Next
Set WS_Printers = Nothing
Set WshNet = Nothing
Set WshShell = Nothing