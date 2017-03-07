<%
/*
Workfile:      dc065Customise.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Cosmetic customisation file for dc065
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		28/09/01	SYS2564/SYS2752 (child) First version, Client specific cosmetic customisation
TLiu	02/09/05	MAR38 - Changed Labels for Flat No., House Name & No.
*/
%>
<SCRIPT LANGUAGE="JavaScript">
function Customise()
{
	document.all("idPostcode").innerHTML = "Postcode";
	document.all("idFlatNo").innerHTML = "Flat No./ Name";
	document.all("idHouseName").innerHTML = "House<br>No. & Name";
	document.all("idDistrict").innerHTML = "District";
	document.all("idTown").innerHTML = "Town";
	document.all("idCounty").innerHTML = "County";				
}

<% // This method contains any overidden methods the client may want to implement. %>
function DerivedScreen()
{
	this.inheritFrom = BaseScreen;
	this.inheritFrom();
}
</SCRIPT>

