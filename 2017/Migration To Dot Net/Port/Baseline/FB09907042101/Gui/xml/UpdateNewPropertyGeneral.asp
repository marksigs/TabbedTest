<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);
	var thisApplication = new ActiveXObject("omApp.NewPropertyBO");
	Response.Write(thisApplication.UpdateNewProperty(xmlIn.xml));
%>