<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);
	
	var thisRelease = new ActiveXObject("omPayProc.PaymentProcessingBO");
	Response.Write(thisRelease.omPayProcRequest(xmlIn.xml));
%>
<% /* OMIGA BUILD VERSION 045.02.05.20.00 */ %>
