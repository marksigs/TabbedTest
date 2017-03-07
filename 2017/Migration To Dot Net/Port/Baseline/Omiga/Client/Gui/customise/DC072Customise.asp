<%
/*
Workfile:      dc072Customise.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Cosmetic customisation file for dc072
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
BS		01/03/02	SYS2564/SYS4211 (child) First version, Client specific cosmetic customisation
*/
%>
<SCRIPT LANGUAGE="JScript">
function Customise() 
{
	document.all("idValuerName").innerHTML = "Valuer Name";
	document.all("idValuationDate").innerHTML = "Valuation Date";
	document.all("idValuationAmount").innerHTML = "Valuation Amount";
}

</SCRIPT>
