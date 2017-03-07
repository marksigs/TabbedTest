<%@  Language=JScript %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);
	
	var thisOtherInsCo = new ActiveXObject("omApp.ApplicationBO");
	//Response.ContentType = "text/xml";
	Response.Write(thisOtherInsCo.UpdateOtherInsuranceCompany(xmlIn.xml));
%>
