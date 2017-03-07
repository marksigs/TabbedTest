<SCRIPT LANGUAGE="JavaScript">
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Prog	Date		Defect		Description
SC      08/02/2006  2391 SYS    Tenure field should be mandatory
SD		10/02/2006				Making Type of Property mandatory
SC		06/03/2006	1339 UAT	Default the Instruction Address correct option to Yes
								and then disable the option buttons
INR		20/12/2006 EP2_532 - Need max & min for PrivateOwnership.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
function SetMasks() 
{
	frmScreen.txtGroundRent.setAttribute("filter","[0-9.]");
	frmScreen.txtServiceCharge.setAttribute("filter","[0-9.]");
	frmScreen.txtDistrict.setAttribute("filter", "[a-zA-Z ]");
	frmScreen.txtFlatNo.setAttribute("filter", "[0-9a-zA-Z ]");
	frmScreen.txtHouseName.setAttribute("filter", "[0-9a-zA-Z ]");
	frmScreen.txtHouseNo.setAttribute("filter",  "[0-9a-zA-Z ]");
	frmScreen.txtPostcode.setAttribute("filter", "[0-9a-zA-Z ]");
	frmScreen.txtPostcode.setAttribute("upper", "true");
	frmScreen.txtChiefRent.setAttribute("filter", "[0-9]");
	//frmScreen.txtRentalIncome.setAttribute("filter", "[0-9]"); // EP2_2- Moved to AP211
	frmScreen.txtEstRoadCharge.setAttribute("filter", "[0-9]");
	frmScreen.txtExtFloorArea.setAttribute("filter", "[0-9]");
	frmScreen.txtStreet.setAttribute("filter", "[0-9a-zA-Z ]");
	frmScreen.txtCounty.setAttribute("filter",  "[0-9a-zA-Z ]");
	frmScreen.txtTown.setAttribute("filter",  "[0-9a-zA-Z ]");
	frmScreen.txtUnexpiredLease.setAttribute("filter", "[0-9]");
	frmScreen.txtYearBuilt.setAttribute("filter", "[0-9]");
	<% /* SC 08/02/2006: Defect 2391 Tenure should be mandatory */ %>
	frmScreen.cboTenure.setAttribute("required", "true");
	frmScreen.cboTypeOfProperty.setAttribute("required", "true");
	// EP2_2 - New field
	frmScreen.txtPrivateOwnership.setAttribute("filter", "[0-9]");
	// EP2_532
	frmScreen.txtPrivateOwnership.setAttribute("integer", "true");
	frmScreen.txtPrivateOwnership.setAttribute("min", "0");
	frmScreen.txtPrivateOwnership.setAttribute("max", "100");
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
        return 0;
}

<% /* Specify client code for populating screen */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
function ClientPopulateScreen() 
{
	//SC 06/03/2006 Defect 1339 Start
	frmScreen.InstructionAddressCorrect_Yes.checked = true;
	frmScreen.InstructionAddressCorrect_Yes.disabled=true;	
	frmScreen.InstructionAddressCorrect_No.disabled=true;	
	//SC 06/03/2006 Defect 1339 End
}
</SCRIPT>