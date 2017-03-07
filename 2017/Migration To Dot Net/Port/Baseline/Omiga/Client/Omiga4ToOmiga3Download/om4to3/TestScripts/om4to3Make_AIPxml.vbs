set obj=CreateObject("om4to3.Omiga4toOmiga3BO")

strXML =  _
	"<REQUEST  USERID=""WIBBLE"" ACTION=""DOWNLOAD"">" & _
		"<TYPE>AIP</TYPE>" & _
		"<APPLICATIONNUMBER>00005010</APPLICATIONNUMBER>" & _
	"</REQUEST>"
t=Timer


Response =  obj.Download(strXML)

msgbox "AIP Time=" & (timer-t) & vbcrlf & response 