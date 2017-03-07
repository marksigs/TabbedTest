<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>
<%
	var xmlIn = new ActiveXObject("Msxml2.FreeThreadedDOMDocument.4.0");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisCustomer = new ActiveXObject("omCust.CustomerBO");
	Response.Write(thisCustomer.GetNumberOfCopiesForKFI(xmlIn.xml));
%>
