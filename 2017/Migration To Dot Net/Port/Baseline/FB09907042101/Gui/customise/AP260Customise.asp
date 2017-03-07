<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      ap260Customise.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Cosmetic customisation file for ap260

History:
Prog	Date		Description
----	----		-----------
AT		25/04/02	SYS4359	- First version, Client specific cosmetic customisation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<SCRIPT LANGUAGE="JScript">
function Customise() 
{
	document.all("idValuationType").innerHTML = "Valuation Type"; 
	document.all("idValuerType").innerHTML = "Valuer Type";
	document.all("idValuationStatus").innerHTML = "Valuation Status";
}
</SCRIPT>
