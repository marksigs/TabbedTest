<%@ Language="JScript" CodePage=65001 %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var obj = new ActiveXObject("omCust.CustomerBO");
	Response.Write(obj.DeleteAlias(xmlIn.xml));
%>
