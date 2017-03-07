<%
/*
Workfile:      dc221Customise.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Cosmetic customisation file for dc222
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MC		25/09/01	SYS2564/SYS2759 (child) First version, Client specific cosmetic customisation
*/
%>
<SCRIPT LANGUAGE="JScript">
function Customise() 
{
	document.all("idOriginalTerm").innerHTML = "Original Term of Lease";
	document.all("idUnexpiredTerm").innerHTML = "Unexpired Term of Lease";
	document.all("idServiceCharges").innerHTML = "Annual Service Charges";
}

</SCRIPT>
