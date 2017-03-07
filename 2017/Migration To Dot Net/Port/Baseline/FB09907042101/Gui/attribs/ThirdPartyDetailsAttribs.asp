<SCRIPT LANGUAGE="JScript">
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history
Prog	Date		AQR		Description
SD		04/10/2005	MAR115	Limit Length of address fields to 30 characters		
SD		05/10/2005	MAR115	Set MaxLength of HouseNUmber = 6, flatNumber=30, HouseName=30, County = 20	
JJ		07/10/2005	MAR119	Validating Flat no., House no. and House name.
							PostCode,Town,Street made mandatory fields.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

function SetThirdPartyDetailsMasks()
{
	frmScreen.txtCompanyName.setAttribute("wildcard", "true");
	frmScreen.txtCompanyName.setAttribute("required", "true");
	frmScreen.txtPostcode.setAttribute("filter", "[0-9a-zA-Z ]");
	frmScreen.txtPostcode.setAttribute("upper", "true");
	//frmScreen.txtTelephoneNo.setAttribute("filter", "[0-9 ]");
	//frmScreen.txtFaxNo.setAttribute("filter", "[0-9 ]");
	//frmScreen.txtEMailAddress.setAttribute("filter", "[-A-Za-z0-9@._/ ]");
	if (m_ctrSortCode != null)
	{
		m_ctrSortCode.setAttribute("mask","99-99-99");
		m_ctrSortCode.setAttribute("required", "true");
		m_ctrSortCode.setAttribute("filter","[0-9- ]");
	}
	if (m_ctrOrganisationType != null)
		m_ctrOrganisationType.setAttribute("required", "true");

	//START: MAR36 - New code added by Joyce Joseph on 10-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.txtHouseName.style.textTransform = "capitalize";
	frmScreen.txtStreet.style.textTransform = "capitalize";
	frmScreen.txtDistrict.style.textTransform = "capitalize";
	frmScreen.txtTown.style.textTransform = "capitalize";
	frmScreen.txtCounty.style.textTransform = "capitalize";
	frmScreen.txtContactForename.style.textTransform = "capitalize";
	frmScreen.txtContactSurname.style.textTransform = "capitalize";
	//END: MAR36
	
	//SD Start-MAR115
	frmScreen.txtHouseNumber.maxLength = 6;
	frmScreen.txtFlatNumber.maxLength = 30;
	frmScreen.txtHouseName.maxLength = 30;
	frmScreen.txtCounty.maxLength = 20;
	//SD End - MAR115
	
	//MAR119 - DRC backed out as this affects several thirdparty screens
	//START: MAR119 - New code added by Joyce Joseph on 07-Oct-2005
	//Made mandatory fields
	//frmScreen.txtTown.setAttribute("required", "true")
	//frmScreen.txtStreet.setAttribute("required", "true")
	//frmScreen.txtPostcode.setAttribute("required", "true")
	//END: MAR119
}


function ThirdPartyDetailsScreenRules()
{
    //MAR119 - DRC backed out as this affects several thirdparty screens
	//START: MAR119 - New code added by Joyce Joseph on 07-Oct-2005
	//Validating Flat no., House no. and House name.
	//if ((frmScreen.txtFlatNumber.value == "") && (frmScreen.txtHouseNumber.value == "") && (frmScreen.txtHouseName.value == ""))
	//{
	//	frmScreen.txtFlatNumber.focus();
	//	alert("Please enter a value for Flat Number or House Number or House name");
	//	return 2;
	//}
	//else
	//{
	//	return 0;
	//}
	//END: MAR119 
}

</SCRIPT>
