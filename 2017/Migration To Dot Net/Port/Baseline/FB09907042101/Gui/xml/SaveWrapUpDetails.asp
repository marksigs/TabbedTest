<%@ Language="JScript" CodePage=65001 %>
<% /* Added 14/10/2005*/ %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisCustomer = new ActiveXObject("omCust.CustomerBO");
	Response.Write(thisCustomer.SaveWrapUpDetails(xmlIn.xml));
%>
