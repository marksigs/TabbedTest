<SCRIPT LANGUAGE="JScript">

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history
Prog	Date		AQR		Description
JJ      21/09/2005  MAR72   Validating Flat no., House no. and House name.  
JJ		27/09/2005  MAR84   Disable Other Forenames on Screens			
SD		04/10/2005	MAR115	Limit Length of address fields to 30 characters	
SD		05/10/2005	MAR115	Set MaxLength of HouseNUmber = 6, flatNumber=30, HouseName=30, County = 20
SD		19/10/2005	MAR209	Allow only alphabets in name fields and numeric characters in phone fields	
JJ		16/12/2005			Street made un-mandatory.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom Specific history
Prog	Date		AQR		Description
pct		15/03/2006	EP8		BFPO Changes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

function SetMasks(bNewCustomer)
{
<%	// MF 20/07/05 - IA_WP01 %>
	//frmScreen.txtSurname.setAttribute("required", "true");
	//frmScreen.txtFirstForename.setAttribute("required", "true");
<%
	//frmScreen.cboTitle.setAttribute("required", "true");
	//frmScreen.txtDateOfBirth.setAttribute("required", "true");
%>
	frmScreen.txtDateOfBirth.setAttribute("date", "DD/MM/YYYY");
<%	// MF 20/07/05 - IA_WP01 
	//frmScreen.cboGender.setAttribute("required", "true");
%>

	//START: MAR36 - New code added by Joyce Joseph on 10-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.txtSurname.style.textTransform = "capitalize";
	frmScreen.txtFirstForename.style.textTransform = "capitalize";
    frmScreen.txtSecondForename.style.textTransform = "capitalize";
	frmScreen.txtOtherForenames.style.textTransform = "capitalize";
	frmScreen.txtCurrentAddressHouseName.style.textTransform = "capitalize";
	frmScreen.txtCurrentAddressStreet.style.textTransform = "capitalize";
	frmScreen.txtCurrentAddressDistrict.style.textTransform = "capitalize";
	frmScreen.txtCurrentAddressTown.style.textTransform = "capitalize";
	frmScreen.txtCurrentAddressCounty.style.textTransform = "capitalize";
	frmScreen.txtCorrespondenceAddressHouseName.style.textTransform = "capitalize";
	frmScreen.txtCorrespondenceAddressStreet.style.textTransform = "capitalize";
	frmScreen.txtCorrespondenceAddressDistrict.style.textTransform = "capitalize";
	frmScreen.txtCorrespondenceAddressTown.style.textTransform = "capitalize";
	frmScreen.txtCorrespondenceAddressCounty.style.textTransform = "capitalize";
	//END: MAR36


	frmScreen.txtCurrentAddressPostcode.setAttribute("upper", "true");
	frmScreen.txtCurrentAddressPostcode.setAttribute("filter", "[0-9a-zA-Z ]");
	frmScreen.txtCorrespondenceAddressPostcode.setAttribute("filter", "[0-9a-zA-Z ]");
	frmScreen.txtCorrespondenceAddressPostcode.setAttribute("upper", "true");
<%	// APS 09/09/99 - UNIT TEST REF 15, 16 ,17
%>	frmScreen.txtContactEMailAddress.setAttribute("filter", "[-A-Za-z0-9@._/ ]");
	frmScreen.txtTelNumber1.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtTime1.setAttribute("filter", "[-.,:0-9a-zA-Z ]");
	frmScreen.txtTelNumber2.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtTime2.setAttribute("filter", "[-.,:0-9a-zA-Z ]");
	frmScreen.txtTelNumber3.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtTime3.setAttribute("filter", "[-.,:0-9a-zA-Z ]");
	frmScreen.txtTelNumber4.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtTime4.setAttribute("filter", "[-.,:0-9a-zA-Z ]");
	frmScreen.txtExtensionNumber1.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtExtensionNumber2.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtExtensionNumber3.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtExtensionNumber4.setAttribute("filter", "[-0-9\(\) ]");
	//JR 04/10/01 - Omiplus24
	frmScreen.txtCountryCode1.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtCountryCode2.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtCountryCode3.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtCountryCode4.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtAreaCode1.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtAreaCode2.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtAreaCode3.setAttribute("filter", "[-0-9\(\) ]");
	frmScreen.txtAreaCode4.setAttribute("filter", "[-0-9\(\) ]");
	//JR - End
	//ASu - BMIDS00106 - Start 
	frmScreen.txtCountryCode1.setAttribute("phone", "[0-9 ]");
	frmScreen.txtCountryCode2.setAttribute("phone", "[0-9 ]");
	frmScreen.txtCountryCode3.setAttribute("phone", "[0-9 ]");
	frmScreen.txtCountryCode4.setAttribute("phone", "[0-9 ]");
	frmScreen.txtAreaCode1.setAttribute("phone", "[0-9 ]");
	frmScreen.txtAreaCode2.setAttribute("phone", "[0-9 ]");
	frmScreen.txtAreaCode3.setAttribute("phone", "[0-9 ]");
	frmScreen.txtAreaCode4.setAttribute("phone", "[0-9 ]");
	frmScreen.txtTelNumber1.setAttribute("phone", "[0-9 ]");
	frmScreen.txtTelNumber2.setAttribute("phone", "[0-9 ]");
	frmScreen.txtTelNumber3.setAttribute("phone", "[0-9 ]");
	frmScreen.txtTelNumber4.setAttribute("phone", "[0-9 ]");
	frmScreen.txtExtensionNumber1.setAttribute("phone", "[0-9 ]");
	frmScreen.txtExtensionNumber2.setAttribute("phone", "[0-9 ]");
	frmScreen.txtExtensionNumber3.setAttribute("phone", "[0-9 ]");
	frmScreen.txtExtensionNumber4.setAttribute("phone", "[0-9 ]");
	//ASu - End
	
	<% /* MF modification to MAR19 */ %>
	if (bNewCustomer){	
		//START: MAR72 - New code added by Joyce Joseph on 20-Sep-2005
		//Made mandatory fields
		frmScreen.txtCurrentAddressTown.setAttribute("required", "true")
		frmScreen.txtCurrentAddressPostcode.setAttribute("required", "true")
		//END: MAR72
	}
	
	//START: MAR84 - New code added by Joyce Joseph on 11-Aug-2005
	// Disable Other Forenames on Screens
	frmScreen.txtOtherForenames.style.visibility="hidden";
	//END: MAR84	
	
	//SD - start MAR115
	frmScreen.txtCurrentAddressHouseNumber.maxLength = 6;
	frmScreen.txtCurrentAddressFlatNumber.maxLength = 30;
	frmScreen.txtCurrentAddressHouseName.maxLength = 30;
	frmScreen.txtCurrentAddressCounty.maxLength = 20;
	frmScreen.txtCorrespondenceAddressHouseNumber.maxLength = 6;
	frmScreen.txtCorrespondenceAddressFlatNumber.maxLength = 30;
	frmScreen.txtCorrespondenceAddressHouseName.maxLength = 30;
	frmScreen.txtCorrespondenceAddressCounty.maxLength = 20;
	//SD - End MAR115
	
	//SD MAR209 Start
	// MAR369 PJO 09/11/2005 - Allow - and '
	frmScreen.txtSurname.setAttribute("filter", "[a-zA-Z '-]");
	frmScreen.txtFirstForename.setAttribute("filter", "[a-zA-Z '-]");
	frmScreen.txtSecondForename.setAttribute("filter", "[a-zA-Z '-]");
	frmScreen.txtOtherForenames.setAttribute("filter", "[a-zA-Z '-]");
	
	frmScreen.txtCountryCode1.setAttribute("filter", "[0-9]");
	frmScreen.txtCountryCode2.setAttribute("filter", "[0-9]");
	frmScreen.txtCountryCode3.setAttribute("filter", "[0-9]");
	frmScreen.txtCountryCode4.setAttribute("filter", "[0-9]");
	frmScreen.txtAreaCode1.setAttribute("filter", "[0-9]");
	frmScreen.txtAreaCode2.setAttribute("filter", "[0-9]");
	frmScreen.txtAreaCode3.setAttribute("filter", "[0-9]");
	frmScreen.txtAreaCode4.setAttribute("filter", "[0-9]");
	frmScreen.txtTelNumber1.setAttribute("filter", "[0-9]");
	frmScreen.txtTelNumber2.setAttribute("filter", "[0-9]");
	frmScreen.txtTelNumber3.setAttribute("filter", "[0-9]");
	frmScreen.txtTelNumber4.setAttribute("filter", "[0-9]");
	frmScreen.txtExtensionNumber1.setAttribute("filter", "[0-9]");
	frmScreen.txtExtensionNumber2.setAttribute("filter", "[0-9]");
	frmScreen.txtExtensionNumber3.setAttribute("filter", "[0-9]");
	frmScreen.txtExtensionNumber4.setAttribute("filter", "[0-9]");
	//SD MAR209 End
}

<% /* Get data required for client validation here */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
function GetRulesData()
{
}

<% /* Specify client validation here */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
function ScreenRules()
{
	<% /* MV - BMIDS00754 - 29/10/2002	
	if(frmScreen.txtContactEMailAddress.value == "")
	{
		frmScreen.txtContactEMailAddress.focus();
		alert("Email address has not been entered");
		return 2;
	}
	else */%>
	

	//START: EP8 pct - Validation for BFPO Addresses
	if (frmScreen.chkCorrespondenceAddressBFPO.checked) {
		if ((frmScreen.txtCorrespondenceAddressTown.value == "") && (frmScreen.txtCorrespondenceAddressDistrict.value == "") && (frmScreen.txtCorrespondenceAddressCounty.value == "")) {
			alert("A BFPO Address must have details in at least one field of: Town, District or County");
			return 2;
		}
		else {
			return 0;
		}
	}
	// EP8 End
	
	//START: (MAR72) - New code added by Joyce Joseph on 21-Sep-2005
	//Validating Flat no., House no. and House name.
	if ((frmScreen.txtCurrentAddressFlatNumber.value == "") && (frmScreen.txtCurrentAddressHouseNumber.value == "") && (frmScreen.txtCurrentAddressHouseName.value == ""))
	{
		frmScreen.txtCurrentAddressHouseNumber.focus();
		alert("Please enter a value for Flat Number or House Number or House name");
		return 2;
	}
	
	if ((frmScreen.txtCorrespondenceAddressPostcode.value != "") || (frmScreen.txtCorrespondenceAddressFlatNumber.value = "") || 	(frmScreen.txtCorrespondenceAddressHouseNumber.value != "") || (frmScreen.txtCorrespondenceAddressHouseName.value != "") ||	(frmScreen.txtCorrespondenceAddressStreet.value != "") ||	(frmScreen.txtCorrespondenceAddressDistrict.value != "") || (frmScreen.txtCorrespondenceAddressTown.value != "") ||	(frmScreen.txtCorrespondenceAddressCounty.value != "") || (frmScreen.cboCorrespondenceAddressCountry.value != ""))	 	
	{
		if ((frmScreen.txtCorrespondenceAddressFlatNumber.value == "") && (frmScreen.txtCorrespondenceAddressHouseNumber.value == "") && (frmScreen.txtCorrespondenceAddressHouseName.value == ""))
		{
			frmScreen.txtCorrespondenceAddressHouseNumber.focus();
			alert("Please enter a value for Flat Number or House Number or House name");
			return 2;
		}
		else if (frmScreen.txtCorrespondenceAddressTown.value == "")
		{
			frmScreen.txtCorrespondenceAddressTown.focus();
			alert("Please enter the Town name");
			return 2;
		}
		else if (frmScreen.txtCorrespondenceAddressPostcode.value == "")
		{
			frmScreen.txtCorrespondenceAddressPostcode.focus();
			alert("Please enter the Postcode");
			return 2;
		}
		else
		{
			return 0;
		}		
	}
	else
	{
		return 0;
	}
	//END: (MAR72)Code added by Joyce Joseph on 21-Sep-2005
}

<% /* Specify client code for populating screen */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
function ClientPopulateScreen() 
{
	<% /* MV - BMIDS00754 - 29/10/2002	
	if(m_sMetaAction == "CreateNewCustomerForNewApplication")
		scScreenFunctions.SetRadioGroupValue(frmScreen, "MailshotInd", "0"); */ %>
}
</SCRIPT>