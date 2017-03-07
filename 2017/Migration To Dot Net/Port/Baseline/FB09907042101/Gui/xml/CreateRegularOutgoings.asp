<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var oApp = new ActiveXObject("omCF.CustomerFinancialBO");
	Response.Write(oApp.CreateRegularOutgoings(xmlIn.xml));
%>
