<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("omCRUD.omCRUDBO");
	Response.Write(thisObject.OmRequest(xmlIn.xml));
%>
