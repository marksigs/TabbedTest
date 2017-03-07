<%@ Language="JScript" %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObj = new ActiveXObject("omCM.MortgageSubQuoteBO");
	Response.Write(thisObj.CreateOneOffCosts(xmlIn.xml));
%>
