<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>
<% /* INR 25/11/2005 MAR645 Created */ %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisCC = new ActiveXObject("omCC.CreditCheckBO");

	Response.Write(thisCC.CreateRuleOverride(xmlIn.xml));
%>
