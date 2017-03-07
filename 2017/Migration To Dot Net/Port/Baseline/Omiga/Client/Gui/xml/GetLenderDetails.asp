<%@ Language="JScript" CodePage=65001 %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObj = new ActiveXObject("omOrg.MortgageLenderBO");
	Response.Write(thisObj.GetLenderDetails(xmlIn.xml));
%>
