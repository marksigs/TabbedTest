<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisApplication = new ActiveXObject("omApp.ApplicationBO");
	Response.Write(thisApplication.CreateApplicationPriority(xmlIn.xml));
%>
