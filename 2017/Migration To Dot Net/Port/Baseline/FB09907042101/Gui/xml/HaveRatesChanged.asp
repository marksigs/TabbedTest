<%@ Language="JScript" CodePage=65001 %>
<%
	// BMIDS624
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObj = new ActiveXObject("omCM.QuotationBO");
	Response.Write(thisObj.HaveRatesChanged(xmlIn.xml));
	// BMIDS624 End
%>
