<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>
<%	
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);
	
	//var fso = new ActiveXObject("Scripting.FileSystemObject");
	//var file = fso.CreateTextFile("c:\\payprocrequest.txt", true);
	//file.WriteLine(xmlIn.xml);
		
	var thisPaymentProcess = new ActiveXObject("omPayProc.PaymentProcessingBO");
	Response.Write(thisPaymentProcess.omPayProcRequest(xmlIn.xml));
%>
