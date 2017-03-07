<%
/*
Workfile:      mc010Customise.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Cosmetic customisation file for mc010
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MC		05/10/01	SYS2564/SYS2775 (child) First version, Client specific cosmetic customisation
*/
%>
<SCRIPT LANGUAGE="JScript">
function Customise() 
{
	document.all("idTerm").innerHTML = "Term";
	document.all("idMonthlyCICost").innerHTML = "Monthly Capital<br>&amp; Interest Cost";
	document.all("spnTerm").style.left = "50px";
}

</SCRIPT>
