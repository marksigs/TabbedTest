<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      ap265Customise.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Cosmetic customisation file for ap265

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
	document.all("idValuationRequired").innerHTML = "Valuation Required"; 
	document.all("idValuationPrice").innerHTML = "Valuation Price";
	document.all("idLastValuer").innerHTML = "Last Valuer";
	document.all("idLastValuationDate").innerHTML = "Last Valuation Date";

	document.all("idPostcode").innerHTML = "Postcode";
	document.all("idFlatNo").innerHTML = "Flat No.";
	document.all("idHouseNameAndNo").innerHTML = "House Name &amp; No.";

	document.all("idDistrict").innerHTML = "District";
	document.all("idTown").innerHTML = "Town";
	document.all("idCounty").innerHTML = "County";

	document.all("idServiceCharge").innerHTML = "Service Charge";
}
</SCRIPT>
