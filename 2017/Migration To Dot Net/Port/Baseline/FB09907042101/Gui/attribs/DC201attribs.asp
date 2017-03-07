<SCRIPT LANGUAGE="JScript">
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history
Prog	Date		AQR		Description
JJ      10/10/2005  MAR119  If  Remortgage, Disable new prop indicator & 
							additional Occupants.
							NumberOfBedrooms only between 1-10
Maha T	28/11/2005	MAR713	Make NumberOfBedrooms mandatory
DRC     01/12/2005  MAR762  Make Tenure mandatory
INR		13/02/2007	EP2_780/788
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

function SetMasks()
{
	frmScreen.txtPurchasePrice.setAttribute("filter", "[0-9]");
	//	AW	06/12/02	BM0116
	frmScreen.txtDiscount.setAttribute("filter", "[0-9]");
	frmScreen.txtSharedAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtPreEmptionDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtPurchasePrice.setAttribute("required", "true");
	<%/* SA 30/5/01 SYS1835 ValuationType & PropertyLocation now mandatory */%>
	<% /* MF 12/08/2005 MAR20 valuation type not always mandatory 
	frmScreen.cboValuationType.setAttribute("required", "true");
	*/ %>
	frmScreen.cboTenureOfProperty.setAttribute("required", "true");
	frmScreen.cboPropertyLocation.setAttribute("required", "true");
	frmScreen.txtDateOfEntry.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtEarliestRemortgage.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtNumberOfBedrooms.setAttribute("filter", "[0-9]");
	
	<% /* MAR713 - Maha T */ %>
	frmScreen.txtNumberOfBedrooms.setAttribute("required", "true");
	
	//START: MAR119 Code changed by Joyce Joseph on 10-Oct-05
	//NumberOfBedrooms only between 1-10
	frmScreen.txtNumberOfBedrooms.setAttribute("integer","true");
	frmScreen.txtNumberOfBedrooms.setAttribute("min","1");
	frmScreen.txtNumberOfBedrooms.setAttribute("max","10");
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
	//START: MAR119 Code added by Joyce Joseph on 10-Oct-2005 
	//If  Remortgage, Disable new prop indicator & additional Occupants.
	if (frmScreen.txtTypeOfApplication.value == "Remortgage")
	{
		frmScreen.optNewPropertyNo.checked=true;
		frmScreen.optOtherResidentsNo.checked=true;
		frmScreen.optNewPropertyYes.disabled=true;	
		frmScreen.optNewPropertyNo.disabled=true;	

		frmScreen.optOtherResidentsYes.disabled=true;	
		frmScreen.optOtherResidentsNo.disabled=true;	
		frmScreen.btnOtherResidents.disabled=true;
	}	
	//END: MAR119
}
</SCRIPT>
