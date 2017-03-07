<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>
<%
	var thisGlobalParameterBO = new ActiveXObject("omBase.GlobalParameterBO");
	Response.ContentType = "text/xml";
	Response.Write(thisGlobalParameterBO.GetCurrentParameter(Request.QueryString("Request")));
%>
