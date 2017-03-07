<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisNewProperty = new ActiveXObject("omApp.NewPropertyBO");
	Response.Write(thisNewProperty.UpdateNewPropertyAddress(xmlIn.xml));
%>
