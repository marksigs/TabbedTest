<%@ Language="JScript" CodePage=65001 %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false
	xmlIn.load(Request);

	var thisParam = new ActiveXObject("omBase.GlobalParameterBO");
	Response.Write(thisParam.GetCurrentParameterListEx(xmlIn.xml));
%>
