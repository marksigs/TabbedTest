<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>
<%
	/* PSC 02/08/2002 BMIDS00006 Created */

	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("omCF.CustomerFinancialBO");
	Response.Write(thisObject.UpdateSpecialFeature(xmlIn.xml));
%>
