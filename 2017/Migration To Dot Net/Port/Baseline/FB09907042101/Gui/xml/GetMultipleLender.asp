<%@  Language=JScript %>
<%
	var thisObject = new ActiveXObject("omBase.GlobalParameterBO");
	Response.Write(thisObject.IsMultipleLender());
%>
