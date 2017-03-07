<%@ Language="JScript" CodePage=65001%>
<%
	var xmlIn = new ActiveXObject("Msxml2.FreeThreadedDOMDocument.4.0");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObj = new ActiveXObject("omAQ.ApplicationQuoteBO");
	Response.Write(thisObj.RemodelQuotation(xmlIn.xml));
%>
