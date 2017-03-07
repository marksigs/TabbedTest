<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisThirdParty = new ActiveXObject("omTP.ThirdPartyBO");
	Response.Write(thisThirdParty.FindPanelValuerList(xmlIn.xml));
%>
