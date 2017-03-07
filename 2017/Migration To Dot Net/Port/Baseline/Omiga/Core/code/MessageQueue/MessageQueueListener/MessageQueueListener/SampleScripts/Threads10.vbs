dim obj
set obj  = createobject("MessageQueueListener.MessageQueueListener1.1")
if not obj is nothing then
	strXML = "<?xml version=""1.0""?>"
	strXML = strXML & "<REQUEST ACTION=""UPDATE"">"
	strXML = strXML & "<QUEUELIST>"
	strXML = strXML & "<QUEUE>"
	strXML = strXML & "<NAME>.\Test</NAME>"
	strXML = strXML & "<TYPE>MSMQ1</TYPE>"
	strXML = strXML & "<TASK>THREADS</TASK>"
	strXML = strXML & "<NUMBER>10</NUMBER>"
	strXML = strXML & "</QUEUE>"
	strXML = strXML & "</QUEUELIST>"
	strXML = strXML & "</REQUEST>"
	MsgBox obj.Configure(strXML),,"Configure Listener"
	set obj = nothing
else
	msmgbox "Could not create object"
end if
