<%
/*
Workfile:      dc182Customise.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Cosmetic customisation file for dc182
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		30/01/02	First version, SYS2564(parent) SYS3959(child) Client specific cosmetic customisation
*/
%>

<SCRIPT LANGUAGE="JScript">

function Customise() 
{
	document.all("idVATNumber").innerHTML = "VAT Number";
	document.all("idPercentageSharesHeld").innerHTML = "Percentage of Shares held";
}

</SCRIPT>
