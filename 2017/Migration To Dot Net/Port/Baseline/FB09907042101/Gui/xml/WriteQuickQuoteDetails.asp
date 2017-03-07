<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("omQQ.QuickQuoteApplicantDetailsBO");
	Response.Write(thisObject.SaveQuickQuoteData(xmlIn.xml));
%>
