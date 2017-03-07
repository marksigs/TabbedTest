<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisAttitudeToBorrowing = new ActiveXObject("omQQ.AttitudeToBorrowingBO");
	Response.Write(thisAttitudeToBorrowing.Create(xmlIn.xml));
%>
