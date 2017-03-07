<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false
	xmlIn.load(Request);

	var thisMortSubQuote = new ActiveXObject("omCM.MortgageSubQuoteBO");
	Response.Write(thisMortSubQuote.ValidateCompulsoryProducts(xmlIn.xml));
%>
