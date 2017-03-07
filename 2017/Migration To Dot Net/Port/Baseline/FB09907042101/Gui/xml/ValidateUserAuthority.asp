<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false
	xmlIn.load(Request);

	var thisAppQuote = new ActiveXObject("OmOrg.OrganisationBO");
	Response.Write(thisAppQuote.ValidateUserAuthority(xmlIn.xml));
%>
