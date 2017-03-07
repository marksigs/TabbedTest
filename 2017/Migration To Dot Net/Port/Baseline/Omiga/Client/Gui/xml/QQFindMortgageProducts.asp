<%@  Language=JScript CodePage=65001 %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("omQQ.QuickQuoteBO");
	Response.Write(thisObject.FindMortgageProducts(xmlIn.xml));
%>
