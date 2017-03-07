<SCRIPT LANGUAGE="JScript">
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history
Prog	Date		AQR		Description
SD		04/10/2005	MAR115	Limit Length of address fields to 30 characters		
SD		05/10/2005	MAR115	Set MaxLength of HouseNUmber = 6, flatNumber=30, HouseName=30, County = 20	
JJ      10/10/2005  MAR119	Postcode,Street,Town made mandatory fields.


~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History

Prog    Date			AQR		Description
pct		07/03/2006		EP211	Postcode Field required - set in main code depending on BFPO Flag
PB		13/04/2006		EP221	Street made non-mandatory field.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
function SetMasks(bPreviousAddress)
{
	if(bPreviousAddress){
		frmScreen.cboApplicantName.setAttribute("required", "true");
		frmScreen.txtPostcode.setAttribute("filter", "[0-9a-zA-Z ]");
		frmScreen.txtPostcode.setAttribute("upper", "true");
		//SYS1007 More mandatory fields
		frmScreen.cboCustomerAddressType.setAttribute("required", "true");
		frmScreen.cboOccupancy.setAttribute("required", "true");		
	}
	frmScreen.txtDateMovedIn.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtDateMovedIn.setAttribute("required", "true");
	frmScreen.txtDateMovedOut.setAttribute("date", "DD/MM/YYYY");
	// ASu - BMIDS00478 Start
	frmScreen.txtPropertyValue.setAttribute("amount","[0-9]");
	frmScreen.txtPropertyValue.setAttribute("filter","[0-9]");
	// ASu - End
	
	//SD Start-MAR115
	frmScreen.txtHouseNumber.maxLength = 6;
	frmScreen.txtHouseName.maxLength = 30;
	frmScreen.txtFlatNo.maxLength = 30;
	frmScreen.txtCounty.maxLength = 20;
	//SD End-MAR115
	
	//START: MAR119 - New code added by Joyce Joseph on 10-Oct-2005
	//Made mandatory fields
	frmScreen.txtTown.setAttribute("required", "true")
	//frmScreen.txtStreet.setAttribute("required", "true") // PB 13/04/2006 EP221
	//frmScreen.txtPostcode.setAttribute("required", "true") // pct 07/03/2006 EP211
	//END: MAR119
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
}

</SCRIPT>