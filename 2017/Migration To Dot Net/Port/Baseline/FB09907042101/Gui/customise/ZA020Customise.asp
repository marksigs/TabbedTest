<%
/*
Workfile:      za020Customise.asp
Copyright:     Copyright © 2001 Marlborough Stirling

Description:   Cosmetic customisation file for za020
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MC		08/10/01	First version, SYS2781 Client specific cosmetic customisation
*/
%>
<SCRIPT LANGUAGE="JScript">
function Customise() 
{
	document.all("idSortCode").innerHTML = "Bank Sort Code";
	document.all("idTown").innerHTML = "Town";
}
</SCRIPT>
