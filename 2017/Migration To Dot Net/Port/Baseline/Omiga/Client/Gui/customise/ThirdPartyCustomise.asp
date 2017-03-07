<%
/*
Workfile:      ThirdPartyCustomise.asp
Copyright:     Copyright © 2001 Marlborough Stirling

Description:   Cosmetic customisation file for ThirdPartyDetails
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MC		08/10/01	First version, SYS2785 Client specific cosmetic customisation
GD		22/02/02	SYS4138 Allow customisation of surname,forename
TLiu	02/09/2005	MAR38 Changed labels for Flat No., House Name & No.
*/
%>
<SCRIPT LANGUAGE="JScript">

function ThirdPartyCustomise() 
{
	document.all("idPostcode").innerHTML = "Postcode";
	document.all("idFlatNumber").innerHTML = "Flat No./ Name";
	document.all("idHouseName").innerHTML = "Building No. &amp; Name";
	document.all("idDistrict").innerHTML = "District";
	document.all("idTown").innerHTML = "Town";
	document.all("idCounty").innerHTML = "County";
	document.all("idSurname").innerHTML = "Surname";
	document.all("idForename").innerHTML = "Forename";
}

<% // This method contains any overidden methods the client may want to implement. %>
function DerivedScreen(sBaseGroupList)
{
	this.inheritFrom = BaseScreen;
	this.inheritFrom(sBaseGroupList);
}
</SCRIPT>
