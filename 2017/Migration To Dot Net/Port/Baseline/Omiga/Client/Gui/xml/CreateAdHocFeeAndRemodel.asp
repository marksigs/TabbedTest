<%@ Language="JScript" %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObj = new ActiveXObject("omAQ.ApplicationQuoteBO");
	Response.Write(thisObj.CreateAdHocFeeAndRemodel(xmlIn.xml));
%>


