<%@ Language="JScript" CodePage=65001 %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false
	xmlIn.load(Request);

	var thisAppQuote = new ActiveXObject("OmAQ.ApplicationQuoteBO");
	Response.Write(thisAppQuote.ValidateUserMandateLevel(xmlIn.xml));
%>
