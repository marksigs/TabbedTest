<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>
<%
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false;
	xmlIn.load(Request);

	//var thisCustreg = new ActiveXObject("omCR.ApplicationLockBO");
	//Response.Write(thisCustreg.UnlockApplicationAndCustomers(xmlIn.xml));
	
	//SR 09/03/00 Call ApplicationManagerBO.UnlockApplicationAndCustomers
	var thisApplManager = new ActiveXObject("omApp.ApplicationManagerBO");
	Response.Write(thisApplManager.UnlockApplicationAndCustomers(xmlIn.xml));
%>
