<%@ LANGUAGE="JSCRIPT" %>
<%
/* GHun		23/12/2002	BM0213 Converted from VBScript to Javascript so the builder doesn't break it */
Response.Cookies("XMLGUIDebugging") = 1;
if (Request.Cookies("XMLGUIDebugging") == 1)
	Response.Redirect("LogonFrameset.asp");
else 
	Response.write("Cannot set debugging on as I was unable to set up a local cookie on your machine. Please check the Cookies setting of your browser (it should be enabled for debugging to work)");	
%>

