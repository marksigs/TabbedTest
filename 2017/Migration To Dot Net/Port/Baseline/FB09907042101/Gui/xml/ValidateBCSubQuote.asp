<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("omCM.BuildingsAndContentsSubQuoteBO");
	Response.Write(thisObject.ValidateSubQuote(xmlIn.xml));	
%>
