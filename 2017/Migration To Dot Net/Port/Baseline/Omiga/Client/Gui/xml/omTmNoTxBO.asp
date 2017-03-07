<%@ Language="JScript" CodePage=65001 %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObj = new ActiveXObject("omTM.omTmNoTxBO");
	Response.Write(thisObj.OmTmNoTxRequest(xmlIn.xml));
%>
