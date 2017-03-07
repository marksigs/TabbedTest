<%@  Language=JavaScript %>
<?xml version="1.0"?>
<%
	var thisQQ = new ActiveXObject("OmigaIVBS.BSLogOn");
	Response.Write(thisQQ.LogOn(Request.QueryString("UserId"),Request.QueryString("Password")));
%>
