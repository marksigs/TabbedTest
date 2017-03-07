<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("omOrg.OrganisationBO");
	Response.Write(thisObject.ValidateUserLogon(xmlIn.xml));
%>
