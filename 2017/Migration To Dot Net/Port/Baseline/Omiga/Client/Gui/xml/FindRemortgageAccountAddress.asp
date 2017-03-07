<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 15 Oct 2002 10:41:24 */ %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);
	
	var thisCustomerFinancial = new ActiveXObject("omCF.CustomerFinancialBO");
	Response.Write(thisCustomerFinancial.FindRemortgageAccountAddress(xmlIn.xml));
%>
