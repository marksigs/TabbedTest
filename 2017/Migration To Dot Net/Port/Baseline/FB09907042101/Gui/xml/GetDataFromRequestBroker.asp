<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 15 Oct 2002 10:41:24 */ %>

<% /* Created by automated update TW 15 Oct 2002 SYS5115 */ %>
<%

	var xmlIn = new ActiveXObject("microsoft.xmldom");
	xmlIn.async = false
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("omRB.OmRequestDO");
	Response.Write(thisObject.OmDataRequest(xmlIn.xml));
	
%>
