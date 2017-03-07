<%@ Language="JScript" CodePage=65001 %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("omAPP.NewPropertyBO");
	Response.Write(thisObject.SaveNewPropertyAndUpdateFactFind(xmlIn.xml));
%>
