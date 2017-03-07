dim obj
set obj  = createobject("MessageQueueListener.MessageQueueListener1.1")
if not obj is nothing then
	strXML = "<?xml version=""1.0""?>"
	strXML = strXML & "<REQUEST ACTION=""GET"">"
	strXML = strXML & "<QUEUELIST>"
	strXML = strXML & "</QUEUELIST>"
	strXML = strXML & "</REQUEST>"
	MsgBox obj.Configure(strXML),,"Configure Listener"
	set obj = nothing
else
	msmgbox "Could not create object"
end if
