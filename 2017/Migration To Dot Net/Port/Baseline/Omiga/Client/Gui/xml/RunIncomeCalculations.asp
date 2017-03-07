<%@ Language="JScript" CodePage=65001 %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObj = new ActiveXObject("omIC.IncomeCalcsBO");
	Response.Write(thisObj.RunIncomeCalculation(xmlIn.xml));
%>

