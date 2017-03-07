<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisCustomerRelationship = new ActiveXObject("omApp.ApplicationBO");
	Response.Write(thisCustomerRelationship.GetCustomerRelationship(xmlIn.xml));
%>
