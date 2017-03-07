<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("omAQ.ApplicationQuoteBO");
	Response.Write(thisObject.CalcAllowableInc(xmlIn.xml));
%>
