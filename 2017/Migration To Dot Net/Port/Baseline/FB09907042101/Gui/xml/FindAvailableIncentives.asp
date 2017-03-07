<%@ Language="JScript" CodePage=65001 %>
<% /* Added to BMids on 17/10/2002 for CPWP1 - DPF */ %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisApp = new ActiveXObject("omCM.MortgageSubQuoteBO");
	Response.Write(thisApp.FindAvailableIncentives(xmlIn.xml));
	
%>