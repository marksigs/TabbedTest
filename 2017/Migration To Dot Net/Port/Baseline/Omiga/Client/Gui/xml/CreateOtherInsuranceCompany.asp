<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);
	
	var thisOtherInsCo = new ActiveXObject("omApp.ApplicationBO");
	//Response.ContentType = "text/xml";
	Response.Write(thisOtherInsCo.CreateOtherInsuranceCompany(xmlIn.xml));
%>
