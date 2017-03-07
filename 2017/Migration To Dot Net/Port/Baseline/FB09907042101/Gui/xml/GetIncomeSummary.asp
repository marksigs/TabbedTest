<%@ Language="JScript" CodePage=65001 %>
<% /* cddd
MV - 18/10/2005 - MAR228
*/ %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisIncomeData = new ActiveXObject("omCE.CustomerEmploymentBO");
	Response.Write(thisIncomeData.GetIncomeSummary(xmlIn.xml));
%>