<%
/*
Workfile:      mc030Customise.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Cosmetic customisation file for mc030
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MC		05/10/01	SYS2564/SYS2776 (child) First version, Client specific cosmetic customisation
*/
%>
<SCRIPT LANGUAGE="JScript">
function Customise() 
{
	document.all("idTerm").innerHTML = "Term";
	document.all("idTerm").style.width = "60px";
	document.all("idTerm").style.textAlign = "right";
	document.all("idCIAmount").innerHTML = "Capital &amp; Interest <br>Amount";
	document.all("idCIMonthlyCost").innerHTML = "Monthly Capital<br>&amp; Interest Cost";
}

</SCRIPT>
