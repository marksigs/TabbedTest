<%@  Language=JScript %>
<% /* MAR1885  Change to use msxml version 4 to correct loading of document contents */ %>
<%
	var xmlIn = new ActiveXObject("Msxml2.FreeThreadedDOMDocument.4.0");
	xmlIn.setProperty("NewParser", true);
	xmlIn.async = false;

	if (xmlIn.load(Request))
	{
		var thisApp = new ActiveXObject("omPrint.omPrintBO");
		Response.Write(thisApp.omRequest(xmlIn.xml));
	}
	else
	{
		var xmlOut = "<RESPONSE TYPE=\"APPERR\">";
		xmlOut += "<ERROR>"
		xmlOut += "<SOURCE>PrintManager.asp</SOURCE>"
		xmlOut += "<DESCRIPTION>Error in loading Request</DESCRIPTION>"
		xmlOut += "</ERROR>"
		xmlOut += "</RESPONSE>";
		Response.Write(xmlOut);
	}
	
%>