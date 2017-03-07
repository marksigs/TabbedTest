<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisPropertyInsDetails = new ActiveXObject("omCust.CustomerBO");
	Response.Write(thisPropertyInsDetails.UpdatePropertyInsuranceDetails(xmlIn.xml));
%>
