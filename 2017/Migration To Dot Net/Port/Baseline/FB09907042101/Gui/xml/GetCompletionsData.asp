<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>
<%

	//Added for Completions - SYS3949
	
	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("omRB.OmRequestDO");
	Response.Write(thisObject.OmDataRequest(xmlIn.xml));
	
	
%>
 <% /* OMIGA BUILD VERSION 028.02.04.17.00 */ %>
