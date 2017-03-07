set objLockManager = CreateObject("omLockManager.LockManagerBO")

requestApplication = _
	"<REQUEST " & _ 
		"USERID=""MSGUser"" " & _
		"UNITID=""10"" " & _
		"MACHINEID=""CH007770"" " & _
		"CHANNELID=""10"" " & _
		"ENABLELOCKING=""1"" " & _
		"CREATENEW=""0"" " & _
		"APPLICATIONNUMBER=""100000002"" " & _
		"OPERATION=""LOCK"">" & _
	"</REQUEST>"

requestCustomer = _
	"<REQUEST " & _ 
		"USERID=""MSGUser"" " & _
		"UNITID=""10"" " & _
		"MACHINEID=""CH007770"" " & _
		"CHANNELID=""10"" " & _
		"CUSTOMERNUMBER=""100000541"" " & _
		"OPERATION=""LOCK"">" & _
	"</REQUEST>"

request = requestApplication	
'request = requestCustomer

if not objLockManager is nothing then
	response = objLockManager.omRequest(request)
	MsgBox response
end if

'		"APPLICATIONNUMBER=""100000703"" " & _
'		"ENABLELOCKING=""1"" " & _
