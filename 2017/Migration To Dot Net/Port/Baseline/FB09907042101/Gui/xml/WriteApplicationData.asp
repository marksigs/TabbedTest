<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisApplicationData = new ActiveXObject("omAPP.ApplicationBO");
	Response.Write(thisApplicationData.UpdateApplication(xmlIn.xml));
%>
