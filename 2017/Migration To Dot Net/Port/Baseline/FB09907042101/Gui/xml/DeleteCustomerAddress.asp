<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisCustomer = new ActiveXObject("omCust.CustomerBO");
	Response.Write(thisCustomer.DeleteCustomerAddress(xmlIn.xml));
%>
