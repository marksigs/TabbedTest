<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      ap200Customise.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Cosmetic customisation file for ap200

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
	//scScrFunctions.SizeTextToField(lblFW030Title,"Valuation Report - Valuation");
	document.all("idValuation").innerHTML = "Valuation"; 
	document.all("btnValuerDetails").value = "Valuer Details"; 
	//document.all("idValuerInvoiceAmount").innerHTML = "Valuer Invoice Amount";
	//document.all("idValuationOK").innerHTML = "Valuation OK";
}
</SCRIPT>
