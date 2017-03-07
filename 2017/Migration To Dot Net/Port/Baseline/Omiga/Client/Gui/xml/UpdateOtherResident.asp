<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisApp = new ActiveXObject("omApp.ApplicationBO");
	Response.Write(thisApp.UpdateOtherResident(xmlIn.xml));
%>
