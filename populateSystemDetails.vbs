startUp = "Analysis Started...This might take few Minutes"
WScript.Echo startUp

'----------------------------------------------------------------------------------------------
'To get the AssetID
'----------------------------------------------------------------------------------------------
strDetails=""
assetId=""
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
For Each objComputer in colSettings 
assetId=objComputer.Name
strDetails = strDetails & "AssetID:" & objComputer.Name & "||"
Next



'----------------------------------------------------------------------------------------------
'To get the HardDisk Size
'----------------------------------------------------------------------------------------------
Const GB = 1073741824
HD=0
networkHD=0
TotalHardDiskSize=0
TotalHD=0
TotalNHD=0
TotalRHD=0
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
Set colDisks = objWMIService.ExecQuery _
    ("Select * from Win32_LogicalDisk")
For Each objDisk in colDisks   
if IsNumeric(objDisk.Size) Then
	if objDisk.DriveType = 3  Then        'For Local hardDisk 
		HD=(HD + (objDisk.Size/GB))
	End If
	
	if objDisk.DriveType = 4 Then         'For Network Disk
		networkHD=(networkHD + (objDisk.Size/GB))
	End If 
	
	if objDisk.DriveType = 2 Then          'For Removable Disk
		removableHD=(removableHD + (objDisk.Size/GB))
	End If
End If
Next

TotalHD=FormatNumber(HD,0)
strDetails = strDetails & "LocalHardDiskSize:" & FormatNumber(HD,0) & "||"


TotalNHD=FormatNumber(networkHD,0)
strDetails = strDetails & "NetworkHardDiskSize:" & FormatNumber(networkHD,0) & "||"


TotalRHD=FormatNumber(removableHD,0)
strDetails = strDetails & "RemovableHardDiskSize:" & FormatNumber(removableHD,0) & "||"


WScript.Echo "20% Completed..."

'----------------------------------------------------------------------------------------------
'To get the Total RAM installed
'----------------------------------------------------------------------------------------------
strComputer = "."
RAM = 0
TRAM=0
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory",,48)
For Each objItem in colItems
RAM = RAM + objItem.Capacity
Next

TRAM=FormatNumber(RAM/GB,0)
strDetails = strDetails & "TotalRamInstalled:" & FormatNumber(RAM/GB,0) & "||"

'----------------------------------------------------------------------------------------------
'To get the Usable RAM memory
'----------------------------------------------------------------------------------------------
strComputer = "."
UsableRam = 0
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
For Each objComputer in colSettings 

strDetails = strDetails & "UsableRam:" & FormatNumber((objComputer.TotalPhysicalMemory)/GB,2) & "||"
Next



'----------------------------------------------------------------------------------------------
'To get the IPAddress
'----------------------------------------------------------------------------------------------
strComputer = "."
IPAddress = null
Set objWMIService = GetObject( _ 
    "winmgmts:\\" & strComputer & "\root\cimv2")

Set IPConfigSet = objWMIService.ExecQuery _
    ("Select IPAddress from Win32_NetworkAdapterConfiguration ")
 
For Each IPConfig in IPConfigSet
    If Not IsNull(IPConfig.IPAddress) Then        

IPAddress=IPConfig.IPAddress(i)
strDetails = strDetails & "IPAddress:" & IPConfig.IPAddress(i) & "||"
    End If
Next

WScript.Echo "40% Completed..."

'----------------------------------------------------------------------------------------------
'To get the OS Details
'----------------------------------------------------------------------------------------------

strComputer = "."
OSVersion = null
OSServicePack = null
OSManufacturer = null
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colSettings 
	
strDetails = strDetails & "Version:" & objOperatingSystem.Version & "||"
	
strDetails = strDetails & "ServicePack:" & objOperatingSystem.ServicePackMajorVersion & "." & objOperatingSystem.ServicePackMinorVersion & "||"
	
strDetails = strDetails & "Manufacturer:" & objOperatingSystem.Manufacturer & "||"
	
strDetails = strDetails & "WindowsDirectory:" & objOperatingSystem.WindowsDirectory & "||"
Next


'----------------------------------------------------------------------------------------------
'To get the OS Name
'----------------------------------------------------------------------------------------------

strComputer = "."
OSName=null
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colOperatingSystems
    
OSName=Replace(objOperatingSystem.Caption," ","_") 
strDetails = strDetails & "OSName:" & objOperatingSystem.Caption  & "||"
Next

'----------------------------------------------------------------------------------------------
'To get the Login Name(Associate Id)
'----------------------------------------------------------------------------------------------

strComputer = "." 
EmpId=0
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_ComputerSystem",,48)
For Each objItem in colItems 

EmpId=Replace(objItem.UserName,"INDIA\","")

strDetails = strDetails & "AssociateId:" & Replace(objItem.UserName,"INDIA\","")  & "||"

Next



WScript.Echo "60% Completed..."

'----------------------------------------------------------------------------------------------
'To get the Associate Name
'----------------------------------------------------------------------------------------------
EmpName=null
Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
strFullName = objUser.Get("displayName")

EmpName=Replace(strFullName," ","_")
strDetails = strDetails & "AssociateName:" & strFullName  & "||"


'----------------------------------------------------------------------------------------------
'To get Name Manufacturer and OS Type(32 or 64 bit)
'----------------------------------------------------------------------------------------------

strComputer = "."
ProcessorType=null
OStype=null
Set objWMIService = GetObject(_
    "winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery(_
    "Select * from Win32_Processor")
For Each objItem in colItems

strDetails = strDetails & "ProcessorManufacturer:" & objItem.Manufacturer  & "||"

ProcessorType=Replace(objItem.Name," ","_")

strDetails = strDetails & "ProcessorType:" & objItem.Name  & "||"

OStype=objItem.AddressWidth & "_bit"
strDetails = strDetails & "OSType:" & objItem.AddressWidth & " bit"  & "||"
Next



WScript.Echo "80% Completed..."
'----------------------------------------------------------------------------------------------
'To get the login users history
'----------------------------------------------------------------------------------------------
showfolderlist "c:\Users"

Sub ShowFolderList(folderspec)
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.SubFolders
    For Each f1 in fc
	if IsNumeric(f1.name)=true then
	s = s + 1
	
strDetails = strDetails & "User" & s & ":" & f1.name  & "||"
	End If
    Next
End Sub


complete = "100% Completed...Analysis Complete"
WScript.Echo complete

Dim wsh
Set wsh=WScript.CreateObject("WScript.Shell")
