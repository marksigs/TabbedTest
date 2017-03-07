<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>
<%
	var thisCustomer = new ActiveXObject("omCust.CustomerBO");
	Response.ContentType = "text/xml";
	Response.Write(thisCustomer.GetCustomerPersonalDetails(Request.QueryString("Request")));
%>