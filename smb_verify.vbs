Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

letter = "v:"

MsgBox "To jest test"

Set colDisks = objWMIService.ExecQuery( "Select * from Win32_LogicalDisk Where Name = """ & letter & """" )
For Each objDisk In colDisks
	drivetype = objDisk.DriveType
	currentmapping = objDisk.ProviderName
Next

MsgBox currentmapping