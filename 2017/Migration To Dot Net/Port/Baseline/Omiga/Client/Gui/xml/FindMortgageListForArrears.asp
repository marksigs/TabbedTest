<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisCustomerFinancial = new ActiveXObject("omCF.CustomerFinancialBO");
	Response.Write(thisCustomerFinancial.FindMortgageListForArrears(xmlIn.xml));
%>
