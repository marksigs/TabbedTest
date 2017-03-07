<%@ Language="JScript" CodePage=65001 %>
<% /* MAR487 file was missing from SourceSafe */ %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);
	
	var thisObj = new ActiveXObject("omApp.ApplicationBO");
	Response.Write(thisObj.GetApplicationUnderwriting(xmlIn.xml));
%>
