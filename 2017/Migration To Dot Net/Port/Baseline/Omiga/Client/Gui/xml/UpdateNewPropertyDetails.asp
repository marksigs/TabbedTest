<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("omApp.NewPropertyBO");
	Response.Write(thisObject.UpdateNewPropertyDetails(xmlIn.xml));
%>
