<%@ Language="JScript" CodePage=65001 %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false
	xmlIn.load(Request);

	var thisDirectoryList = new ActiveXObject("omTP.ThirdPartyBO");
	Response.Write(thisDirectoryList.GetDirectoryDetails(xmlIn.xml));
%>
