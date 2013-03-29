Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

letter = "v:"

MsgBox "To jest test"

' It's a little bit of comment

Set colDisks = objWMIService.ExecQuery( "Select * from Win32_LogicalDisk Where Name = """ & letter & """" )
For Each objDisk In colDisks
	drivetype = objDisk.DriveType
	currentmapping = objDisk.ProviderName
Next

MsgBox currentmapping