<%@  Language=JScript %>
<% /* MAR645 Created */ %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisAPP = new ActiveXObject("omAPP.ApplicationBO");
	Response.Write(thisAPP.SaveApplicationUnderwriting(xmlIn.xml));
%>
