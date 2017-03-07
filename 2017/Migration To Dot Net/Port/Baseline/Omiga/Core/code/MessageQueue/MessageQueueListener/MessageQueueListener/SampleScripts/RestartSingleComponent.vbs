dim obj
set obj  = createobject("MessageQueueListener.MessageQueueListener1.1")
if not obj is nothing then
	strXML = "<?xml version=""1.0""?>"
	strXML = strXML & "<REQUEST ACTION=""UPDATE"">"
	strXML = strXML & "<QUEUELIST>"
	strXML = strXML & "<QUEUE>"
	strXML = strXML & "<NAME>.\andy</NAME>"
	strXML = strXML & "<TYPE>MSMQ1</TYPE>"
	strXML = strXML & "<TASK>RESTARTSINGLECOMPONENT</TASK>"
	strXML = strXML & "<COMPONENTLIST>"
	strXML = strXML & "<COMPONENT>MessageQueueComponentVC.MessageQueueComponentVC1</COMPONENT>"
	strXML = strXML & "</COMPONENTLIST>"
	strXML = strXML & "</QUEUE>"
	strXML = strXML & "</QUEUELIST>"
	strXML = strXML & "</REQUEST>"
	MsgBox obj.Configure(strXML),,"Configure Listener"
	set obj = nothing
else
	msmgbox "Could not create object"
end if
