<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("omCM.MortgageSubQuoteBO");
	Response.Write(thisObject.Update(xmlIn.xml));
%>
