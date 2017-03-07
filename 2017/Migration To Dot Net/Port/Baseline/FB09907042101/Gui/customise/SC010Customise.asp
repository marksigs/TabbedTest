<%
/*
Workfile:      sc010Customise.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Cosmetic customisation file for sc010
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MDC		16/04/2002	SYS4396		Client specific cosmetic customisation
LD		19/10/2005  MAR123		NT Authentication
*/
%>
<SCRIPT LANGUAGE="JScript">
function Customise() 
{
	var XML = new scXMLFunctions.XMLObject();
	var strSecurityCredentialsType = XML.GetGlobalParameterString(document, "SecurityCredentialsType");
	if (strSecurityCredentialsType == "WindowsAuthentication")
	{
		document.all("idPassword").innerHTML = "NT Password";
		document.all("btnChangePassword").disabled = true;
	}
	else
	{
		document.all("idPassword").innerHTML = "Omiga Password";
	}
}

</SCRIPT>
