<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("omCM.PaymentProtectionSubQuoteBO");
	Response.Write(thisObject.ValidateSubQuote(xmlIn.xml));	
%>
