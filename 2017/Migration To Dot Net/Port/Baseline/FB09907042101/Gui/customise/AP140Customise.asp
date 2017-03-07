<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      ap140Customise.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Cosmetic customisation file for ap140

History:
Prog	Date		Description
----	----		-----------
AT		10/04/02	SYS4359	- First version, Client specific cosmetic customisation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<SCRIPT LANGUAGE="JScript">
function Customise() 
{
	document.all("idPostcode").innerHTML = "Postcode"; 
	document.all("idBuildingName").innerHTML = "Building Name & No";
	document.all("idDistrict").innerHTML = "District";
	document.all("idTown").innerHTML = "Town";
	document.all("idCounty").innerHTML = "County";
}
</SCRIPT>
