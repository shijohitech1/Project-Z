dim row 'as integer
''''''''''''''''''''''''''''''''''''''''''''''Memory'''''''''''''''''''''''''''''''''''''

dim ram2, ramtype, ramsize(6), rams
dim t

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")
rams =""
For Each objItem in colItems
'ram 4294967296
 for t=0 to 3
	Wscript.Echo "Manufacturer: " & objItem.Manufacturer
	Wscript.Echo "Memory Type: " & objItem.MemoryType
		If objItem.MemoryType = 24 Then
		ramtype = "DDR3"
		ElseIf objItem.MemoryType = 0 Then
		ramtype = "DDR4"
		Else
		ramtype = "DDR2"
		End If
    Wscript.Echo "Capacity: " & objItem.Capacity
		if objItem.Capacity > 4000000000 and objItem.Capacity < 5294967296 Then
		ramsize(t) = "4 GB "
		rams = rams & " " & ramsize(t)
		ElseIf objItem.Capacity > 8000000000 and objItem.Capacity < 5294967296 Then
		ramsize(t) = "8 GB "
		Else 
		ramsize(t) = objItem.Capacity
		end if
    'Wscript.Echo "Data Width: " & objItem.DataWidth
   ' Wscript.Echo "Description: " & objItem.Description
   ' Wscript.Echo "Device Locator: " & objItem.DeviceLocator
    Wscript.Echo "Form Factor: " & objItem.FormFactor
		If objItem.FormFactor = 8 Then
		systemtype = "Laptop"
		ElseIf objItem.FormFactor = 16 Then
		systemtype = "Desktop"
		Else
		systemtype = "Other"
		End If
    'Wscript.Echo "Hot Swappable: " & objItem.HotSwappable
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
	
   ' Wscript.Echo "Memory Type: " & objItem.MemoryType
    'Wscript.Echo "Name: " & objItem.Name
    ''Wscript.Echo "Part Number: " & objItem.PartNumber
    'Wscript.Echo "Position In Row: " & objItem.PositionInRow
    'Wscript.Echo "Speed: " & objItem.Speed
	ram2=objItem.Manufacturer & " " & ramtype
	t=t+1
	next
Next

'rams = ramsize(0) & "+" & ramsize(1) & ramsize(2) & ramsize(3)


'''''''''''''''''''''''''''''''''''''''''''''''Operating System''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
Dim os
'strComputer = "."
'Set objWMIService = GetObject("winmgmts:" _
'& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colOperatingSystems = objWMIService.ExecQuery _
("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colOperatingSystems
'Wscript.Echo "Caption: " & objOperatingSystem.Caption
os=objOperatingSystem.Caption
'Wscript.Echo "Serial Number: " & objOperatingSystem.SerialNumber
'Wscript.Echo "Version: " & objOperatingSystem.Version
Next



'''''''''''''''''''''''''''''''''''''''''''''''User Name''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set wshShell = CreateObject( "WScript.Shell" )
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )



''''''''''''''''''''''''''''''''''IP Address & MAC Address'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim NIC1, Nic, StrIP, CompName, macIP

Set NIC1 =     GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

For Each Nic in NIC1

    if Nic.IPEnabled then
        StrIP = Nic.IPAddress(0)
		macIP = Nic.MACAddress(0)
        'Set WshNetwork = WScript.CreateObject("WScript.Network")
        'CompName= WshNetwork.Computername
       ' MsgBox "IP Address:  "&StrIP & vbNewLine _
        '    & "Computer Name:  "&CompName,4160,"IP Address and Computer Name"
        'wscript.quit
    End if
Next


'''''''''''''''''''''''''''''''''''''HostName''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set wshShell = CreateObject( "WScript.Shell" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
'WScript.Echo "Computer Name: " & strComputerName

'''''''''''''''''''''''''''''''''''''''''''''''Excel'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
strExcelPath = "E:\temp\vbsct.xlsx"
Set objExcel = CreateObject("Excel.Application") 
'view the excel program and file, set to false to hide the whole process
objExcel.Visible = True 
Set objWorkbook = objExcel.Workbooks.Open(strExcelPath)
set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
nUsedRows = objSheet.UsedRange.Rows.Count
row=nUsedRows+1
'get a cell value and set it to a variable
'r3c5 = objExcel.Cells(row,3).Value
objExcel.Cells(row,2).Value = strUserName
objExcel.Cells(row,3).Value = strComputerName
objExcel.Cells(row,4).Value = strIP
objExcel.Cells(row,5).Value = macIP
objExcel.Cells(row,6).Value = os
objExcel.Cells(row,8).Value = systemtype
objExcel.Cells(row,9).Value = ram2
objExcel.Cells(row,10).Value = rams
objWorkbook.Save
'objWorkbook.Close 
'objExcel.Quit
'release objects
'Set objExcel = Nothing
'Set objWorkbook = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
