<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false
	xmlIn.load(Request);

	var thisObj = new ActiveXObject("omOrg.OrganisationBO");
	Response.Write(thisObj.ValidateCurrentPassword(xmlIn.xml));
%>

