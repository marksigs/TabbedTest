<%@ Language="JScript" CodePage=65001 %>
<%
	var xmlIn = new ActiveXObject("Msxml2.FreeThreadedDOMDocument.4.0");
	xmlIn.async = false;
	xmlIn.load(Request);

	var thisObject = new ActiveXObject("omPack.PackManagerBO");
	/* MAR792 GHun fixed spelling of OmRequest */
	Response.Write(thisObject.OmRequest(xmlIn.xml));	
%>
