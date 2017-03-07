<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      ap201Customise.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Cosmetic customisation file for ap201

History:
Prog	Date		Description
----	----		-----------
AT		25/04/02	SYS4359	- First version, Client specific cosmetic customisation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom History:
Prog	Date		Description
----	----		-----------
TW      02/02/2007  	EP2_1021 - There is a spelling mistake on one of the fields in the Valuer Details AP201 screen.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<SCRIPT LANGUAGE="JScript">
function Customise() 
{
	document.all("idForename").innerHTML = "Forename"; 
	document.all("idSurname").innerHTML = "Surname";
	document.all("idPostcode").innerHTML = "Postcode";
	document.all("idBuildingNameNo").innerHTML = "Building&nbsp;Name&nbsp;&&nbsp;No.";
	document.all("idDistrict").innerHTML = "District";
	document.all("idTown").innerHTML = "Town";
	document.all("idCounty").innerHTML = "County";
	document.all("idValuerType").innerHTML = "Valuer&nbsp;Type";
	document.all("idValuationType").innerHTML = "Valuation&nbsp;Type";
}
</SCRIPT>
