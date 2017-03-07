<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisQuotation = new ActiveXObject("omCM.QuotationBO");
	Response.Write(thisQuotation.Update(xmlIn.xml));
%>
