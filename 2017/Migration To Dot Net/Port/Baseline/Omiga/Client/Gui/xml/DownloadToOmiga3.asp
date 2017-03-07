<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisDL = new ActiveXObject("om4to3.Omiga4toOmiga3BO");
	Response.Write(thisDL.Download(xmlIn.xml));
%>
