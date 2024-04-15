Option Explicit
Dim strServer, arrTest
strServer = "."

arrTest = BIOS(strServer)

Function BIOS(strComputer)
'Retrieves BIOS information for a computer, including BIOS version number and release date.
Dim objWMIService, colBIOS, objBIOS, arrBIOS(15,0), x, y, arrBIOSChar(0,0)
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colBIOS = objWMIService.ExecQuery ("Select * from Win32_BIOS")
For each objBIOS in colBIOS
    arrBIOS(1,x) =  "Build Number: " & objBIOS.BuildNumber
    arrBIOS(1,x) =  "Current Language: " & objBIOS.CurrentLanguage
    arrBIOS(1,x) =  "Installable Languages: " & objBIOS.InstallableLanguages
    arrBIOS(1,x) =  "Manufacturer: " & objBIOS.Manufacturer
    arrBIOS(1,x) =  "Name: " & objBIOS.Name
    arrBIOS(1,x) =  "Primary BIOS: " & objBIOS.PrimaryBIOS
    arrBIOS(1,x) =  "Release Date: " & objBIOS.ReleaseDate
    arrBIOS(1,x) =  "Serial Number: " & objBIOS.SerialNumber
    arrBIOS(1,x) =  "SMBIOS Version: " & objBIOS.SMBIOSBIOSVersion
    arrBIOS(1,x) =  "SMBIOS Major Version: " & objBIOS.SMBIOSMajorVersion
    arrBIOS(1,x) =  "SMBIOS Minor Version: " & objBIOS.SMBIOSMinorVersion
    arrBIOS(1,x) =  "SMBIOS Present: " & objBIOS.SMBIOSPresent
    arrBIOS(1,x) =  "Status: " & objBIOS.Status
    arrBIOS(1,x) =  "Version: " & objBIOS.Version
    For i = 0 to Ubound(objBIOS.BiosCharacteristics)
        arrBIOSChar(0,y) =  "BIOS Characteristics: " & objBIOS.BiosCharacteristics(i)
    Next
Next
BIOS= arrBIOS
End Function