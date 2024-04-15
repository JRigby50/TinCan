'  This script, which is made up of routines mostly taken from the MS Scripting Center, creates an html page
'  giving the following information on selected machines
'  Responding to a ping (in RED if it is not responding)  *** Note that ping only works from XP or 2003 machines
'                                                             so the script must either be run on an XP machine
'                                                             OR modified so the pingit function hosts on an XP 
'                                                             machine somewhere in your network.
'  Uptime (in RED if the uptime is over 45 days)
'  if a web server, web sites pinging in black, not pinging in RED
'  if a print server, printer status, with printer status in RED if there is a message (like jammed, off line or toner low)
'  Physical drive status, with drives not "ok" in RED (like mirrors resynching, etc).
'  Logical drives on the physical drive with drives having less than 25% disk space free showing in RED
'  Selected services running, with any "Auto" services stopped showing in RED, "Disabled" services running showing in RED
'  and "Manual" services running showing, but not in red.

'  Constants, not in use, but which I've included.

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

'  Replace my machine names with your own.  I include all servers and my management console (workstation)
strMachines = "Huygens-s;Hubble-s;JBayer;J-BODE;J-Oort;Nikola-Tesla;Hipparchus-S;J-Flamsteed;Aristarchus-s;g-bruno-s"
Set objFSO = CreateObject("Scripting.FileSystemObject")

'  This is the web page file we will be creating (or overwriting if it already exists).  
'  Change to where you want the page.

Set objTextFile = objFSO.CreateTextFile("v:\server.html", True)
 

' Set the Browser Title and Header Information
sWebWriteHeader("SALC-CS Hourly Server Status")

' Print the date and time to the web page (it uses the date/time on the machine from which the script is run).
strDisplay=fGetDate(".")
sWebWriteLine strDisplay

sWebWriteHeading "Current Server Status","",24.0,TRUE,True

aMachines = split(strMachines, ";")
For Each machine in aMachines
   SectionCounter=SectionCounter + 1
   strComputer = Machine
'  Print the Machine Name

   select Case ucase(strComputer)
'     This section replace the actual machine name with the machine's function, on the web page.  Mainly to slow down
'     hackers, as with the machine name and this information, it would be easier to know what machines to target.
      case "EVERGREEN"
         strDisplayName="Development SQL Server"
      case "J-FLAMSTEED"
         strDisplayName="Development SQL Server"
      case "HIPPARCHUS-S"
         strDisplayName="Management Console"
      case "NIKOLA-TESLA"
         strDisplayName="Production Print Server"
      case "HUBBLE-S"
         strDisplayName="Production SQL, Mail and DC Server"
      case "HUYGENS-S"
         strDisplayName="Root Domain Controller"
      case "JBAYER"
         strDisplayName="Production Print and DC Server"
      case "J-OORT"
         strDisplayName="Production File Server"
      case "J-BODE"
         strDisplayName="Production Web Server"
      Case "G-BRUNO-S"
         strDisplayName="Production DHCP Server"
      Case "ARISTARCHUS-S"
         strDisplayName="Future Root Domain Controller"
   end select

   if not PingIt(strComputer) then
      ' Determine if the machine is up, if not, we can not report any other status info
      strDisplay = strDisplayName & " not responding to Ping" 
      sComputerIsUp = false
   else
      sComputerIsUp = true
   end if

   if sComputerIsUp then   
      '  The machine is up
      Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
      Set colSettings = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
      For Each objOperatingSystem in colSettings 
'        The print queue information does not display for windows 2000, but does for XP (and 2003).
         if not instr(objOperatingSystem.Name,"XP")>0 and not instr(objOperatingSystem.Name,"XP") > 0 then
           Not2000 = false
         else
           Not2000 = true           
         end if
	 if instr(objOperatingSystem.Name,"2003")>0 then
           IS2003 = true
         else
           IS2003 = false
         end if
         WhatOS=left(objOperatingSystem.Name,Instr(objOperatingSystem.Name, "|")-1)
      next
      set objWMIService = nothing
      Set colSettings = nothing


      Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
      '  Get the computer name from the current system, we're still using the original objWMIService call from above
      Set colSettings = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
      For Each objComputer in colSettings 
      '  We could be using the passed parameter, but what if no parameter was passed?
         strComputer=objComputer.Name
         strMakeModel=""
'        Make and model, if available
         if not objComputer.Manufacturer="" then
            strMakeModel=strMakeModel & "System Manufacturer: " & objComputer.Manufacturer & ", "
         end if 
         if not objComputer.Model="" then
            strMakeModel=strMakeModel & "System Model: " & objComputer.Model         
         end if
         intMB=int(int(objComputer.TotalPhysicalMemory/1000000)/16)*16
         strRam="Total RAM: " & intMB & " MB"
      Next
      set objWMIService = nothing
      Set colSettings = nothing

      sWebOpenSection(SectionCounter)
      sWebWriteHeading "Machine Name: ",strDisplayName,"12.0",False,True
      ' make and model if available
      sWebWriteBoldLine "Make and/or Model"
      if not strMakeModel="" then
             sWebWriteIndentLine strMakeModel
             sWebWriteIndentLine WhatOS
      else
             sWebWriteIndentLine "Not available from system"
             sWebWriteIndentLine WhatOS
      end if 

      sWebWriteBoldLine "Processors"
      Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
      Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
      For Each objItem in colItems
         countProc=countProc + 1
         ProcessorStr=""
         ProcessorStr=ProcessorStr & "Device ID: " & objItem.DeviceID 
         ProcessorStr=ProcessorStr & ", " & ltrim(objItem.Name)
         ProcessorStr=ProcessorStr & ", " & objItem.Manufacturer
         StatusProc=""
         select Case objItem.StatusInfo
            Case 1
                StatusProc = "Other"
            Case 2
                StatusProc = "Unknown"
            Case 3
                StatusProc = "Enabled"
            Case 4
                StatusProc = "Disabled"
            Case 5
                StatusProc = "N/A"
         end select
         ProcessorStr=ProcessorStr & ", Status: " & StatusProc
         if objItem.CurrentClockSpeed > 0 and not isnull(objItem.CurrentClockSpeed) then
            intGB=round(objItem.CurrentClockSpeed  /1000,1)
            ProcessorStr=ProcessorStr & ", Speed: " & intGB & " Ghz"
         else
            ProcessorStr=ProcessorStr & ", Speed: " & objItem.CurrentClockSpeed & " Ghz"
         end if
         ProcessorStr=ProcessorStr & " " & ", Address Width: " & objItem.AddressWidth & " Bit"
         ProcessorStr=ProcessorStr & " " & ", Data Width: " & objItem.DataWidth & " Bit"
         if Instr(objItem.Name, "Xeon") and not countProc/2 = int(countProc/2) then
             ' Do Nothing, as Xeon's report twice as many processors than they have
         else
             if objItem.StatusInfo = 3 then
                sWebWriteIndentLine ProcessorStr
             else
                sWebWriteIndentRedLine ProcessorStr
             end if  
         end if
      Next
      countProc=0
      set objWMIService = nothing
      Set colItems = nothing

        
      Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
      Set colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
      For Each objOS in colOperatingSystems              
          sWebWriteBoldLine "Memory (RAM and Disk) Status"
          intMB2=int(int(objOS.FreePhysicalMemory/1000))
          AvailPhysmemstr="Available Physical Memory: " & intMB2 & " MB - "
          if intMB2/intMB < .05 then
             ' Less than 5% of RAM is free.
             sWebWriteIndentRedLine strRam & ", " & AvailPhysmemstr & int((intMB2/intMB)*100) & "% Free" 
          else
             sWebWriteIndentLine strRam & ", " & AvailPhysmemstr & int((intMB2/intMB)*100) & "% Free" 
          end if
          intMB=int(int(objOS.TotalVirtualMemorySize/1000))
          TotVirtmemstr= "Total Virtual Memory: " & intMB & " MB"
          intMB2=int(int(objOS.FreeVirtualMemory/1000))
          AvailVirtmemstr = "Available Virtual Memory: " &  intMB2 & " MB - "
          if intMB2/intMB < .05 then         
             ' Less than 5% of swap file is free.
             sWebWriteIndentRedLine TotVirtmemstr & ", " & AvailVirtmemstr & int((intMB2/intMB)*100) & "% Free" 
          else
             sWebWriteIndentLine TotVirtmemstr & ", " & AvailVirtmemstr & int((intMB2/intMB)*100) & "% Free" 
          end if
          intMB=int(int(objOS.SizeStoredInPagingFiles/1000))
          Pagefile= "Page File Space: " & intMB & " MB" 
          sWebWriteIndentLine Pagefile
'         Determine how long since the last reboot.  Windows machines like to be rebooted at least every 45 days.  A planned
'         reboot is better than a mystery crash due to memory leaks or unreleased resources over time.
          dtmBootup = objOS.LastBootUpTime
          dtmLastBootupTime = WMIDateStringToDate(dtmBootup)
          dtmSystemUptime = DateDiff("h", dtmLastBootUpTime, Now)
          sWebWriteBoldLine "Uptime Status"
          if int(dtmSystemUptime/24) > 45 then
             sWebWriteIndentRedLine strDisplayName & " has been running for " & int(dtmSystemUptime/24) & " Days since Reboot"
          else
            sWebWriteIndentLine strDisplayName & " has been running for " & int(dtmSystemUptime/24) & " Days since Reboot"
          end if
      Next
      set objWMIService = nothing
      Set colOperatingSystems = nothing


'     We report special things on Web and Print servers

      select Case machine
'        This is a web server, so we want to ping the web sites to give us an idea if they are up.
         case "J-BODE"
             sWebWriteBoldLine "WEB Site Status"
             if not PingIt("www.salc.wsu.edu") then
                sWebWriteIndentRedLine "salc" & "  is not responding"
             else
                sWebWriteIndentLine "salc" & "  is responding"
             end if
             if not PingIt("www.careers.wsu.edu") then
                sWebWriteIndentRedLine "Careers" & "  is not responding"
             else
                sWebWriteIndentLine "Careers" & "  is responding"
             end if
             if not PingIt("www.sssp.wsu.edu") then
                sWebWriteIndentRedLine "SSSP" & "  is not responding"
             else
                sWebWriteIndentLine "SSSP" & "  is responding"
             end if
             if not PingIt("www.counsel.wsu.edu") then
                sWebWriteIndentRedLine "Counsel" & "  is not responding"
             else
                sWebWriteIndentLine "Counsel" & "  is responding"
             end if
             if not PingIt("www.adcaps.wsu.edu") then
                sWebWriteIndentRedLine "adcaps" & "  is not responding"
             else
                sWebWriteIndentLine "adcaps" & "  is responding"
             end if
             if not PingIt("www.palouseastro.wsu.edu") then
                sWebWriteIndentRedLine "Palouse Astro" & "  is not responding"
             else
                sWebWriteIndentLine "Palouse Astro" & "  is responding"
             end if
             if not PingIt("schedulesurfer.wsu.edu") then
                sWebWriteIndentRedLine "Schedulesurfer" & "  is not responding"
             else
                sWebWriteIndentLine "Schedulesurfer" & "  is responding"
             end if
             if not PingIt("www.techsupport.wsu.edu") then
                sWebWriteIndentRedLine "Techsupport" & "  is not responding - Bet you NEVER see this..."
             else
                sWebWriteIndentLine "Techsupport" & "  is responding"
             end if
'        This is a web server, so we want to ping the web sites to give us an idea if they are up.
         case "Nikola-Tesla"
             sWebWriteBoldLine "WEB Site Status"
             if not PingIt("salctest1.salc.wsu.edu") then
                sWebWriteIndentRedLine "Salctest1" & "  is not responding"
             else
                sWebWriteIndentLine "Salctest1" & "  is responding"
             end if
             if not PingIt("salctest2.salc.wsu.edu") then
                sWebWriteIndentRedLine "Salctest2" & "  is not responding"
             else
                sWebWriteIndentLine "Salctest2" & "  is responding"
             end if
             if not PingIt("salctest3.salc.wsu.edu") then
                sWebWriteIndentRedLine "Salctest3" & "  is not responding"
             else
                sWebWriteIndentLine "Salctest3" & "  is responding"
             end if
             if not PingIt("salctest4.salc.wsu.edu") then
                sWebWriteIndentRedLine "Salctest4" & "  is not responding"
             else
                sWebWriteIndentLine "Salctest4" & "  is responding"
             end if
             if not PingIt("schedulesurfer2.salc.wsu.edu") then
                sWebWriteIndentRedLine "Schedulesurfer2" & "  is not responding"
             else
                sWebWriteIndentLine "Schedulesurfer2" & "  is responding"
             end if
      end select

      select Case machine
         case "JBayer","Nikola-Tesla"
'        This is a print server, so we want to check the shares and status to give us an idea if they are up.
             sWebWriteBoldLine "Printer Status" 
             Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
             Set colInstalledPrinters = objWMIService.ExecQuery ("Select * from Win32_Printer")
             For Each objPrinter in colInstalledPrinters
                Select Case objPrinter.PrinterStatus
                   Case 1
                      strPrinterStatus = "Other"
                   Case 2
                      strPrinterStatus = "Unknown"
                   Case 3
                      strPrinterStatus = "Idle"
                   Case 4
                      strPrinterStatus = "Printing"
                   Case 5
                      strPrinterStatus = "Warm-up"
                End Select
                tmpString = "Name: " & objPrinter.Name & " Location: " & objPrinter.Location & " "
'                sWebWriteIndentLine "Name: " & objPrinter.Name & " Location: " & objPrinter.Location
                if isnull(objPrinter.ServerName) then
'                  if there is not ServerName, then we ARE the server
                   strPrinterName="Local"
                else
                   strPrinterName=objPrinter.ServerName 
                end if
                IF strPrinterStatus = "Other" OR strPrinterStatus = "Unknown" or (objPrinter.Availability<>3 and objPrinter.Availability<>15) Then            
'                  This can be anything from "Toner low" to "Error"
                   sWebWriteIndentRedLine tmpString & " Server Name: " & strPrinterName & ", Share Name: " & objPrinter.ShareName & ", Printer Status: " & strPrinterStatus
                Else
                   sWebWriteIndentLine tmpString & "Server Name: " & strPrinterName & ", Share Name: " & objPrinter.ShareName & ", Printer Status: " & strPrinterStatus
                end if
             Next
             Set objWMIService = Nothing
             Set colInstalledPrinters = Nothing
'            For Windows XP or 2003 servers, you can do this.
             if Not2000=True then 
                Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
                Set colPrintQueues = objWMIService.ExecQuery ("Select * from Win32_PerfFormattedData_Spooler_PrintQueue Where Name <> '_Total'")
                sWebWriteBoldLine "Print Queue Status" 
                if not isnull(colPrintQueues) then                
                   For Each objPrintQueue in colPrintQueues
                      sWebWriteIndentLine "Name: " & objPrintQueue.Name & " Current jobs: " & objPrintQueue.Jobs
                   Next  
               end if
             end if
             Set objWMIService = Nothing
             Set colPrintQueues = Nothing

          case else
'            If it isn't a web server or a print server, it will come here.  I keep this for testing purposes.             
'            wscript.echo machine
      end select


      sWebWriteLine ""

      sWebWriteBoldLine "Drive Status"
      Set wmiServices = GetObject ("winmgmts:{impersonationLevel=Impersonate}!//" & strComputer)
      Set wmiDiskDrives = wmiServices.ExecQuery ("SELECT *, DeviceID FROM Win32_DiskDrive")
      For Each wmiDiskDrive In wmiDiskDrives
'        Once for each physical drive in the server
         if wmiDiskDrive.Status="OK" then
            sWebWriteLine "Physcial Drive: " & wmiDiskDrive.Caption & " Status: " & wmiDiskDrive.Status & " Size: " & round((wmiDiskDrive.Size/1048576)/1000,2) & " GBs"
         else
            sWebWriteRedLine "Physcial Drive: " & wmiDiskDrive.Caption & " Status: " & wmiDiskDrive.Status & " Size: " & round((wmiDiskDrive.Size/1048576)/1000,2) & " GBs"
         end if
         strEscapedDeviceID = Replace(wmiDiskDrive.DeviceID, "\", "\\", 1, -1, vbTextCompare)
         Set wmiDiskPartitions = wmiServices.ExecQuery ("ASSOCIATORS OF {Win32_DiskDrive.DeviceID=""" & strEscapedDeviceID & """} WHERE AssocClass = Win32_DiskDriveToDiskPartition")
         For Each wmiDiskPartition In wmiDiskPartitions
'            Once for each partition on the physical drive.  I keep the tmpDisplay for debugging, but I didn't find the info  meaninful
             tmpDisplay= wmiDiskPartition.DeviceID & " "
             tmpBlockSize=wmiDiskPartition.BlockSize
             if wmiDiskPartition.Bootable then
                tmpDisplay=tmpDisplay & " Bootable: " & wmiDiskPartition.Bootable
             end if
'             sWebWriteLine tmpDisplay
'            We do want the PARTITION SIZE, as opposed to the physical drive size.
             tmpNumberofBlocks=wmiDiskPartition.NumberOfBlocks
             tmpSize=round((tmpBlockSize*tmpNumberOfBlocks/1048576)/1000,2)
             Set wmiLogicalDisks = wmiServices.ExecQuery ("ASSOCIATORS OF {Win32_DiskPartition.DeviceID=""" & wmiDiskPartition.DeviceID & """} WHERE AssocClass = Win32_LogicalDiskToPartition")
             For Each wmiLogicalDisk In wmiLogicalDisks
'              Once for each logical drive (C, D, E, etc.).
               tmpFree=round((wmiLogicalDisk.FreeSpace/1048576)/1000,2)
               tmpPercent=round(tmpFree/tmpSize*100,0)
'              I consider it important if the drive has less than 25% free space, so it will print in Red.
               if IS2003 then  ' The WMI Volume object is new for 2003, and the partition space changed with 2003.
                  Set wmiVolume = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
                  Set colItems = wmiVolume.ExecQuery("Select * from Win32_Volume")
                  For Each objItem In colItems
                     if wmiLogicalDisk.DeviceID = objItem.DriveLetter then
                        tmpfree = round((objItem.FreeSpace/1048576)/1000,2)
                        tmpSize = round((objItem.Capacity/1048576)/1000,2)
                        tmpPercent=round(tmpFree/tmpSize*100,0)
   		        if objItem.Compressed then
                           Compressed=" DoubleSpaced "
                        else
                           Compressed=""
                        end if
		        if objItem.IndexingEnabled then
                           Indexing=" Context Indexing Enabled "
                        else
                           Indexing=""
                        end if
		        MaximumFileNameLength=", Max Filename=" & objItem.MaximumFileNameLength
		        FileSystem=", " & objItem.FileSystem
                     end if
                  Next    
                  if tmpPercent<25 and not IS2003 then                   
                     SWebWriteIndentRedLine "Logical Drive: " & wmiLogicalDisk.DeviceID & Compressed & IndexingEnabled & " " & wmiLogicalDisk.VolumeName & FileSystem & MaximumFileNameLength & ", " & "Diskspace " & tmpFree & " of " & tmpSize & " GBs Free" & " (" & tmpPercent & "%)" &vbCRLF
                  else
                     sWebWriteIndentLine "Logical Drive: " & wmiLogicalDisk.DeviceID & Compressed & IndexingEnabled & " " &  wmiLogicalDisk.VolumeName & FileSystem & MaximumFileNameLength & ", " & "Diskspace " & tmpFree & " of " & tmpSize & " GBs Free" & " (" & tmpPercent & "%)" &vbCRLF
                  end if
               else
                  if tmpPercent<25 and not IS2003 then                  
                     sWebWriteIndentRedLine "Logical Drive: " & wmiLogicalDisk.DeviceID & " " & wmiLogicalDisk.VolumeName & " " & tmpFree & " of " & tmpSize & " GBs Free" & " (" & tmpPercent & "%)" &vbCRLF
                  else
                     sWebWriteIndentLine "Logical Drive: " & wmiLogicalDisk.DeviceID & " " & wmiLogicalDisk.VolumeName & " " & tmpFree & " of " & tmpSize & " GBs Free" & " (" & tmpPercent & "%)" &vbCRLF
                  end if
               end if
             Next ' Drive      
         Next ' next partition
      Next  'next drive       
      set wmiServices = nothing
      Set wmiDiskDrives = nothing


      sWebWriteLine ""
         sWebWriteBoldLine "Services Running on " & strDisplayName
      sWebWriteLine ""
      ' create the table
      WebOpenTable
      ' Create the first row containing the header
      MyColor = "Teal"
      sWriteHeaderRow MyColor,"75","130.95","Name"
      sWriteHeaderRow MyColor,"68","126.2","Description"
      sWriteHeaderRow MyColor,"96","146.8","Start Mode"
      sWriteHeaderRow MyColor,"37","127.45","Service State"
      sWriteHeaderRow MyColor,"49","36.9","*Alert*"
      objTextFile.WriteLine "</tr>"

'     Now what services are on the computer
      Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
      Set colRunningServices = objWMIService.ExecQuery ("Select * from Win32_Service")
'     These are on the computer
      For Each objService in colRunningServices 
'        These are the properties that can be checked on each service running or stopped
'        objService.SystemName, objService.Name, objService.ServiceType, objService.State, 
'        objService.ExitCode, objService.ProcessID, objService.AcceptPause, objService.AcceptStop
'        objService.Caption, objService.Description, objService.DesktopInteract, objService.DisplayName
'        objService.ErrorControl, objService.PathName, objService.Started, objService.StartMode
'        objService.StartName
         sReportit = False
'        There are only certain ones that interest me, generally, so I set the sReportit flag on those services
         select case objService.DisplayName
           case "Alerter"
              sReportit = True
           case "APC PBE Agent"
              sReportit = True
           case "Automatic Updates"
              sReportit = True
           case "Background Intelligent Transfer Service"  
'             Backup Exec uses this and it MUST be running to backup
              sReportit = True
           case "Backup Exec Remote Agent for Windows NT/2000"
              sReportit = True
           case "Backup Exec 8.x Agent Browser"
              sReportit = True
           case "Backup Exec 8.x Alert Server"
              sReportit = True
           case "Backup Exec 8.x Device & Media Service"
              sReportit = True
           case "Backup Exec 8.x Notification Server"
              sReportit = True
           case "Backup Exec 8.x"
              sReportit = True
           case "Backup Exec 8.x"
              sReportit = True
           case "Backup Exec 8.x"
              sReportit = True
           case "Computer Browser"
              sReportit = True
           case "Logical Disk Manager"
              sReportit = True
           case "Distributed File System"
              sReportit = True
           case "DHCP Client"
              sReportit = True
           case "DNS Server"
              sReportit = True
           case "DNS Client"
              sReportit = True
           case "Eudora WorldMail Ph2Ldap Proxy"
              sReportit = True
           case "Event Log"
              sReportit = True
           case "Eudora WorldMail Directory Service"
              sReportit = True
           case "Eudora WorldMail"
              sReportit = True
           case "Eudora WorldMail WEB/LDAP/X.500 Gateway"
              sReportit = True
           case "File Replication Service"
              sReportit = True
           case "File Server for Macintosh"
              sReportit = True
           case "IIS Admin Service"
              sReportit = True
           case "Intel Alert Originator"
              sReportit = True
           case "IPSEC Policy Agent"
              sReportit = True
           case "Kerberos Key Distribution Center"
              sReportit = True
           case "Messenger"
              sReportit = True
           case "MSSQLSERVER"
              sReportit = True
           case "Net Logon"
              sReportit = True
           case "Print Spooler"
              sReportit = True
           case "Remote Registry Service"
              sReportit = True
'          case "Remote Procedure Call (RPC) Locator"
'             sReportit = True
           case "Remote Procedure Call (RPC)"
              sReportit = True
           case "ScheduleSurfer"
              sReportit = True
           case "Security Accounts Manager"
              sReportit = True
           case "Server"
              sReportit = True
           case "Simple Mail Transport Protocol (SMTP)"
              sReportit = True
           case "SPC PBE Server"
              sReportit = True
           case "SQLSERVERAGENT"
              sReportit = True
           case "Symantec AntiVirus Server"
              sReportit = True
           case "System Event Notification"
              sReportit = True
           case "System Event Notification"
              sReportit = True
           case "TCP/IP Print Server"
              sReportit = True
           case "Telnet"
              sReportit = True
           case "Terminal Services"
              sReportit = True
           case "Windows Time"
              sReportit = True
           case "WMDM PMSP"
              sReportit = True
           case "Workstation"
              sReportit = True
           case "World Wide Web Publishing"
              sReportit = True
           case else 
'              If you uncomment the next line, you see all the services, running or stopped.  However, that is too much info.
              sReportit = True
         end select    
         if objService.StartMode = "Auto" AND objService.State = "Stopped" then
'            Services that should be running automatically but are stopped must be experiencing difficulty.  Show in Red.
             sReportit=True
             sFlagIt=True
         end if
         if objService.StartMode = "Disabled" AND objService.State = "Running" then
'            Services that are disabled but are running must be experiencing difficulty.  Show in Red. 
'            In theory, you will probably never see any that meet this criteria.
             sReportit=True
             sFlagIt=True
         end if
         if objService.StartMode = "Manual" AND objService.State = "Running" then
'            Services that are set to run Manual but are running are either called by another program 
'           (and should probably be changed to automatic, not manual) or an indication of hacking activity.  
'           I'm not fully convinced yet whether to flag these in Red.
             sReportit=True
'             sFlagIt=True
         end if
'        The following two if statements turn off any stopped service that is not configured to run.
'        You could comment them out for a full report
         if objService.StartMode = "Manual" AND objService.State = "Stopped" then
             sReportit=False
             sFlagIt=False
         end if
         if objService.StartMode = "Disabled" AND objService.State = "Stopped" then
             sReportit=False
             sFlagIt=False
         end if
         if sReportit then
'           Ok, so we can add a new Table Data cell
            MyCounter=MyCounter+1
'           And color in every other row
            if MyCounter/2 = int(MyCounter/2) then
               if sFlagIt=True then
'                 One shade of red for even rows
                  MyColor="FF6699"
               else
                  MyColor="CCFFFF"
               end if
            else
               if sFlagIt=True then
'                 A different shade of red for odd rows.
                  MyColor="FF3366"
               else
                  MyColor="White"
               end if
            end if
            objTextFile.WriteLine "<tr style='mso-yfti-irow:" & MyCounter & "'>"
            sWriteDataRow MyColor,"75","130.95",objService.Name
            sWriteDataRow MyColor,"68","126.2", objService.Description
            sWriteDataRow MyColor,"96","146.8", objService.StartMode
            sWriteDataRow MyColor,"37","27.45", objService.State
            if sFlagIt=True then
               sWriteDataRow MyColor,"49","36.9", "*Alert*"
            else   
               sWriteDataRow MyColor,"49","36.9", "OK"
            end if
            objTextFile.WriteLine "</tr>"
         end if
         sReportit=false
         sFlagIt=false
      Next
      set objWMIService = nothing
      set colRunningServices = nothing


'     During the loop, we can't know how many rows there are, so when we finish looping, we need to close the table cells.
'     The final row will not actually display, as null cells do not print, but is necessary to close the table.
      MyCounter=MyCounter+1
      objTextFile.WriteLine "<tr style='mso-yfti-irow:" & MyCounter & "mso-yfti-lastrow:yes'>"
      sWriteDataRow MyColor,"75","130.95"," "
      sWriteDataRow MyColor,"68","126.2"," "
      sWriteDataRow MyColor,"96","146.8"," "
      sWriteDataRow MyColor,"37","27.45"," "
      objTextFile.WriteLine "</tr>"

      ' Close the Table
      objTextFile.WriteLine "</table>"
   else
      sWebOpenSection(SectionCounter)
      '  Print the Machine Name
      sWebWriteHeading "Machine Name: ",strComputer,"12.0",False,True
      sWebWriteRedLine strComputer & " not responding to Ping" 
      objTextFile.WriteLine "<p class=MsoNormal><o:p>&nbsp;</o:p></p>"
   end if
   objTextFile.WriteLine "<p class=MsoNormal><o:p>&nbsp;</o:p></p>"
   objTextFile.WriteLine "</div>"
next
'  this next is for the "next" machine in the array.

'  I found a message at the bottom to let me know the script finished was useful, especially as against my systems the 
'  script takes about 7 minutes to run.
objTextFile.WriteLine "<p class=MsoNormal><o:p>All Done</o:p></p>"
objTextFile.WriteLine "</body>"
'  This, in addition to the opening lines of the first header, are required to ensure the file doesn't cache.
objTextFile.WriteLine "<HEAD>"
objTextFile.WriteLine "<META HTTP-EQUIV=" & chr(34) & "PRAGMA" & chr(34) & " CONTENT=" & chr(34) & "NO-CACHE" & chr(34) & ">"
objTextFile.WriteLine "</HEAD>"
'  Page is completed.
objTextFile.WriteLine "</html>"
'  Close the file
objTextFile.Close
Set objTextFile = Nothing
Set objFSO = Nothing

'Wscript.echo "Done"

 
'  Sub-routines and Functions called by the main program


'  Web functions  necessary to write the web pages

sub sWebWriteHeader(sTitle)
   ' Start writing the HTML page in Word Document Style
   objTextFile.WriteLine "<head>"
   objTextFile.WriteLine "<META HTTP-EQUIV=" & chr(34) & "Pragma" & chr(34) & " CONTENT=" & chr(34) & "no-cache" & chr(34) & ">"
   objTextFile.WriteLine "<META HTTP-EQUIV=" & chr(34) & "Expires" & chr(34) & " CONTENT=" & chr(34) & "-1" & chr(34) & ">"
   objTextFile.WriteLine "<meta http-equiv=Content-Type content=" & chr(34) & "text/html; charset=windows-1252" & chr(34) & ">"
   objTextFile.WriteLine "<meta name=ProgId content=Notepad.exe>"
   objTextFile.WriteLine "<meta name=Generator content=" & chr(34) & "Notepad.exe & chr(34) & "">"
   objTextFile.WriteLine "<meta name=Originator content=" & chr(34) & "Notepad.exe" & chr(34) & ">"
' Set the Browser Title and Header Information
   objTextFile.WriteLine "<title>" & sTitle & " " & "</title>"
   objTextFile.WriteLine "<!--[if gte mso 9]><xml>"
   objTextFile.WriteLine " <o:DocumentProperties>"
'  Give credit where credit is due
   objTextFile.WriteLine "  <o:Author>Karen Gallaghar</o:Author>"
   objTextFile.WriteLine "  <o:LastAuthor>Karen Gallaghar</o:LastAuthor>"
   objTextFile.WriteLine "  <o:Revision>1</o:Revision>"
   objTextFile.WriteLine "  <o:TotalTime>2</o:TotalTime>"
   objTextFile.WriteLine "  <o:Created>2004-02-14T22:09:00Z</o:Created>"
   objTextFile.WriteLine "  <o:LastSaved>2005-11-2T22:11:00Z</o:LastSaved>"
   objTextFile.WriteLine "  <o:Company>SALC-CS²</o:Company>"
   objTextFile.WriteLine " </o:DocumentProperties>"
   objTextFile.WriteLine "</xml><![endif]--><!--[if gte mso 9]><xml>"
   objTextFile.WriteLine "</xml><![endif]-->"
'    Set the style...
   objTextFile.WriteLine "<style>"
   objTextFile.WriteLine "<!--"
   objTextFile.WriteLine " /* Style Definitions */"
   objTextFile.WriteLine " p.MsoNormal, li.MsoNormal, div.MsoNormal"
   objTextFile.WriteLine "      {mso-style-parent:"";"
   objTextFile.WriteLine "      margin:0in;"
   objTextFile.WriteLine "      margin-bottom:.0001pt;"
   objTextFile.WriteLine "      mso-pagination:widow-orphan;"
   objTextFile.WriteLine "      font-size:12.0pt;"
   objTextFile.WriteLine "      font-family:" & chr(34) & "Times New Roman" & chr(34) & ";" & chr(34)
   objTextFile.WriteLine "      mso-fareast-font-family:" & chr(34) & "Times New Roman" & chr(34) & ";}"
   objTextFile.WriteLine "@page Section1"
   objTextFile.WriteLine "      {size:8.5in 11.0in;"
   objTextFile.WriteLine "      margin:1.0in 1.25in 1.0in 1.25in;"
   objTextFile.WriteLine "      mso-header-margin:.5in;"
   objTextFile.WriteLine "      mso-footer-margin:.5in;"
   objTextFile.WriteLine "      mso-paper-source:0;}"
   objTextFile.WriteLine "div.Section1"
   objTextFile.WriteLine "      {page:Section1;}"
   objTextFile.WriteLine "-->"
   objTextFile.WriteLine "</style>"
   objTextFile.WriteLine "<!--[if gte mso 10]>"
   objTextFile.WriteLine "<style>"
   objTextFile.WriteLine " /* Style Definitions */"
   objTextFile.WriteLine " table.MsoNormalTable"
   objTextFile.WriteLine "      {mso-style-name:" & chr(34) & "Table Normal" & chr(34) & ";"
   objTextFile.WriteLine "      mso-tstyle-rowband-size:0;"
   objTextFile.WriteLine "      mso-tstyle-colband-size:0;"
   objTextFile.WriteLine "      mso-style-noshow:yes;"
   objTextFile.WriteLine "      mso-style-parent: " & chr(34) & chr(34) & ";"  & chr(34)
   objTextFile.WriteLine "      mso-padding-alt:0in 5.4pt 0in 5.4pt;"
   objTextFile.WriteLine "      mso-para-margin:0in;"
   objTextFile.WriteLine "      mso-para-margin-bottom:.0001pt;"
   objTextFile.WriteLine "      mso-pagination:widow-orphan;"
   objTextFile.WriteLine "      font-size:10.0pt;"
   objTextFile.WriteLine "      font-family:" & chr(34) & "Times New Roman" & chr(34) & ";}"
   objTextFile.WriteLine "</style>"
   objTextFile.WriteLine "<![endif]-->"
   objTextFile.WriteLine "</head>"
end sub

sub sWebOpenSection(sSection)
   objTextFile.WriteLine "<body lang=EN-US style='tab-interval:.5in'>"
   objTextFile.WriteLine "<div class=Section" & sSection & ">"
end sub

sub sWebWriteHeading(sHeading,sField,sFontSize,sCentered,sBolded)
'   sFontSize should be 12.0, 24.0, etc
'   sCentered should be True or False
'   sBolded should be True or False
   sPrint=""
   if sCentered then
      sPrint=sPrint & "<p class=MsoNormal align=center style='text-align:center'>"
   else
      sPrint=sPrint & "<p class=MsoNormal>"
   end if
   if sBolded then
      sPrint=sPrint &  "<b>"
   else

   end if
      sPrint=sPrint &  "<span style='font-size:" & sFontSize & "pt'>" & sHeading & sField & "<o:p></o:p></span>"
   if sBolded then
      sPrint=sPrint &  "</b>"
   end if
   sPrint=sPrint &  "</p>"
   objTextFile.WriteLine sPrint
   objTextFile.WriteLine "<p class=MsoNormal><o:p>&nbsp;</o:p></p>"
end sub

sub sWebWriteLine (sWhatToSay)
   objTextFile.WriteLine "<p class=MsoNormal><span style='font-size:10.0pt'><o:p>" & sWhatToSay & " </o:p></span></p>"
end sub

sub sWebWriteBoldLine (sWhatToSay)
'   Say it in Bold
    objTextFile.WriteLine "<p class=MsoNormal><span style='font-size:10.0pt'><o:p><B>" & sWhatToSay & " </B></o:p></span></p>"
end sub

sub sWebWriteRedLine (sWhatToSay)
'   Say it in RED
    objTextFile.WriteLine "<span style='font-size:10.0pt;color:red'>" & sWhatToSay & "</span><BR>"
end sub

sub sWebWriteIndentLine (sWhatToSay)
'   Say it indented
    objTextFile.WriteLine "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style='font-size:10.0pt'>" & sWhatToSay & "</span><BR>"
end sub

sub sWebWriteIndentRedLine (sWhatToSay)
'   Say it indented AND in RED
    objTextFile.WriteLine "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style='font-size:10.0pt;color:red'>" & sWhatToSay & "</span><BR>"
end sub

sub WebOpenTable()
' create the table
  sTABLELINE1 = "<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 width=1079"
  sTABLELINE2 = " style='width:700pt;border-collapse:collapse;border:none;mso-border-alt:solid windowtext .5pt;" ' 809.4
  sTABLELINE3 = " mso-yfti-tbllook:480;mso-padding-alt:0in 5.4pt 0in 5.4pt;mso-border-insideh:"
  STABLELINE4  = " .5pt solid windowtext;mso-border-insidev:.5pt solid windowtext'>"
  objTextFile.WriteLine   sTABLELINE1 & sTABLELINE2 & sTABLELINE3 & STABLELINE4 
  objTextFile.WriteLine ""
end sub

sub sWriteHeaderRow(Color,ColumnWidth,OtherWidth,Contents)
   '  This creates the bolded header row for the table
   STARTTABLEFIELDDEF="<td"
   STARTCOLOR=" BGCOLOR="
   STARTWIDTH=" width="
   MIDDLETABLEFIELDDEF=" valign=top style='width:"
   ENDTABLEFIELDDEF = "pt;border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>"
   PARACLASSLINE = "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'>"
   SPANSTYLE="<spanstyle='font-size:8.0pt'>"
   CLOSING ="<o:p></o:p></span></b></p></TD>"
   if Isnull(Color) then
     Color="White"
   else
     Color=Color
   end if
   if Isnull(Contents) or Contents="" then
'    Blank cells are suppressed, so we need something in each cell.
     Contents="N/A"
   else
     Contents=Contents
   end if
   objTextFile.WriteLine STARTTABLEFIELDDEF & STARTCOLOR & Color & STARTWIDTH & ColumnWidth & MIDDLETABLEFIELDDEF & OtherWidth & ENDTABLEFIELDDEF
   objTextFile.WriteLine PARACLASSLINE & SPANSTYLE & Contents & CLOSING
   objTextFile.WriteLine ""
end sub

sub sWriteDataRow(Color,ColumnWidth,OtherWidth,Contents)
   '  This creates the un-bolded data rows for the table
   STARTTABLEFIELDDEF="<td"
   STARTCOLOR=" BGCOLOR="
   STARTWIDTH=" width="
   MIDDLETABLEFIELDDEF=" valign=top style='width:"
   ENDTABLEFIELDDEF = "pt;border:solid windowtext 1.0pt;border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>"
   PARACLASSLINE = "<p class=MsoNormal>"
   SPANSTYLE="<span style='font-size:8.0pt'>"
   CLOSING ="<o:p></o:p></span></p></td>"
   if Isnull(Color) then
     Color="White"
   else
     Color=Color
   end if
   if Isnull(Contents) or Contents="" then
'    Blank cells are suppressed, so we need something in each cell.
     Contents="N/A"
   else
     Contents=Contents
   end if
   objTextFile.WriteLine STARTTABLEFIELDDEF & STARTCOLOR & Color & STARTWIDTH & ColumnWidth & MIDDLETABLEFIELDDEF & OtherWidth & ENDTABLEFIELDDEF
   objTextFile.WriteLine PARACLASSLINE & SPANSTYLE & Contents & CLOSING
   objTextFile.WriteLine ""
   objTextFile.WriteLine ""
end sub

'  Functions not related to web page creation

function fGetdate(fComputerName)
   '  Get the system date and time for the header on the web page
   Set objWMIService = GetObject("winmgmts:\\" & fComputerName & "\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_LocalTime")
   For Each objItem in colItems
'      Calculate the day of the week
       select case objItem.DayOfWeek
         case 0
            fstrDayOWeek = "Sunday" 
         case 1
            fstrDayOWeek = "Monday" 
         case 2
            fstrDayOWeek = "Tuesday" 
         case 3
            fstrDayOWeek = "Wednesday" 
         case 4
            fstrDayOWeek = "Thursday" 
         case 5
            fstrDayOWeek = "Friday" 
         case 6
            fstrDayOWeek = "Saturday" 
         case else
            fstrDayOWeek = "A New Day" & objItem.DayOfWeek
       end select
       fstrDisplay = fstrDisplay & "Run on: " & fstrDayOWeek & ", "
'      Want a 2 digit month, even for January - September
       sMonth=objItem.Month 
       if len(ltrim(sMonth))<2 then
          sMonth="0" & sMonth
       else
          sMonth=sMonth
       end if
       fstrDisplay = fstrDisplay & sMonth & "/" 
'      Want a 2 digit day, even for the 1st through the 9th
       sDay=objItem.Day
       if len(ltrim(sDay))<2 then
          sDay="0" & sDay
       else
          sDay=sDay
       end if
       fstrDisplay = fstrDisplay & sDay & "/" 
       fstrDisplay = fstrDisplay & objItem.Year & " at: " 
'      Want a 2 digit hour, even for 1AM - 9AM
       sHour=objItem.Hour
       if len(ltrim(sHour))<2 then
          sHour="0" & sHour
       else
          sHour=sHour
       end if
'      Want a 2 digit minute, even for the first 9 minutes of the hour
       fstrDisplay = fstrDisplay & sHour & ":" 
       sMinute=objItem.Minute
       if len(ltrim(sMinute))<2 then
          sMinute="0" & sMinute
       else
          sMinute=sMinute
       end if
       fstrDisplay = fstrDisplay & sMinute
'       Keep in code in case you ever need the syntax
'       Wscript.Echo "Quarter: " & objItem.Quarter
'       Wscript.Echo "Second: " & objItem.Second
   Next
   fGetdate=fstrDisplay
   set objItems = nothing
   set colItems = nothing
   set objWMIService = nothing  
end function

Function PingIt(Address)
   Dim sHost            'name of Windows XP computer from which the PING command will be initiated
                        'Ping will only work from XP or 2003 machines.
   Dim sTarget          'name or IP address of remote computer to which connectivity will be tested
   Dim cPingResults     'collection of instances of Win32_PingStatus class
   Dim oPingResult              'single instance of Win32_PingStatus class
   sHost = "."
   sTarget = Address
   if len(sTarget)<2 then
  '    Get the computer name from the current system, we're still using the original objWMIService call from above
      Set objWMIService = GetObject("winmgmts:\\" & sTarget & "\root\cimv2")
      Set colSettings = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
      For Each objComputer in colSettings 
'        We could be using the passed parameter, but what if no parameter was passed?
         sTarget=objComputer.Name
      Next
'     Clean up the objects with which we are done.
      set objComputer = nothing
      set colSettings = nothing
      set objWMIService = nothing  ' actually, we'll be using this again in a second.
   end if
   Set cPingResults = GetObject("winmgmts:{impersonationLevel=impersonate}//" & sHost & "/root/cimv2"). ExecQuery("SELECT * FROM Win32_PingStatus " & "WHERE Address = '" + sTarget + "'")
   For Each oPingResult In cPingResults
      If oPingResult.StatusCode = 0 Then
         If LCase(sTarget) = oPingResult.ProtocolAddress Then
            PingIt=True
         Else
            PingIt=True
         End If
      Else
         PingIt=False
      End If
   Next
   set cPingResults = nothing
   set oPingResult = nothing
end function

Function WMIDateStringToDate(dtmBootup)
    WMIDateStringToDate = CDate(Mid(dtmBootup, 5, 2) & "/" & Mid(dtmBootup, 7, 2) & "/" & Left(dtmBootup, 4) _
         & " " & Mid (dtmBootup, 9, 2) & ":" & Mid(dtmBootup, 11, 2) & ":" & Mid(dtmBootup, 13, 2))
End Function

'  Tha' tha' that's all, folks
'  Karen Gallaghar,  Obsessive/Compulsive recovering programmer