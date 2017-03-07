<%@ Language="JScript" CodePage=65001 %>
<% /* Added 'CodePage=65001' SYS5242 by program 09 Oct 2002 12:10:12 */ %>
<% /* MAR1885  Change to use msxml version 4 */ %>
<%
	var xmlIn = new ActiveXObject("Msxml2.FreeThreadedDOMDocument.4.0");
	xmlIn.setProperty("NewParser", true);

	xmlIn.async = false;
	xmlIn.load(Request);

	//DR Old Stylee
	//var thisObject = new ActiveXObject("PrintManager.PrintManagerBO");
	//Response.Write(thisObject.OmRequest(xmlIn.xml));
	
	//DR New Stylee
	var thisObject = new ActiveXObject("omPM.PrintManagerBO");
	Response.Write(thisObject.OmRequest(xmlIn.xml));	
	//Response.Write(thisObject.OmRequest('<REQUEST OPERATION="GETDOCUMENTARCHIVE" FILEGUID="8391FDB745E611D58256001020ACA9F7" FILEVERSION="V1"/>'));
%>
