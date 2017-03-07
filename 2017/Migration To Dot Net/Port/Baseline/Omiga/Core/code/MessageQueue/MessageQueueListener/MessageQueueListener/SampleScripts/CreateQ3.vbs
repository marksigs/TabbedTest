dim obj
set obj  = createobject("MessageQueueListener.MessageQueueListener1.1")
if not obj is nothing then
	strXML = "<?xml version=""1.0""?>"
	strXML = strXML & "<REQUEST ACTION = ""CREATE"">"
	strXML = strXML & "<QUEUELIST>"
	strXML = strXML & "<QUEUE>"
	strXML = strXML & "<NAME>.\AndyOracle</NAME>"
	strXML = strXML & "<TYPE>OMMQ1</TYPE>"
	strXML = strXML & "<CONNECTIONSTRING>DSN=NrRes;UID=test;PWD=password;</CONNECTIONSTRING>"
	strXML = strXML & "<POLLINTERVAL>1000</POLLINTERVAL>"
	strXML = strXML & "<THREADSLIST>"
	strXML = strXML & "<THREADS>"
	strXML = strXML & "<NUMBER>5</NUMBER>"
	strXML = strXML & "</THREADS>"
	strXML = strXML & "</THREADSLIST>"
	strXML = strXML & "</QUEUE>"
	strXML = strXML & "</QUEUELIST>"
	strXML = strXML & "</REQUEST>"
	MsgBox obj.Configure(strXML),,"Configure Listener"
	set obj = nothing
else
	msmgbox "Could not create object"
end if
