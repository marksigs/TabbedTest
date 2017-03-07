<SCRIPT LANGUAGE="JScript">
//GD	25/06/02	BMIDS0077 Remove refs to ids that have been commented out.
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history
Prog	Date		AQR		Description
JJ      10/10/2005  MAR119  Fields disabled
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

function SetMasks() 
{
	frmScreen.txtCustDateStarted.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtDateStarted.setAttribute("date", "DD/MM/YYYY");
	//frmScreen.txtCustGrossBasic.setAttribute("filter", "[0-9]");
	//frmScreen.txtGrossBasic.setAttribute("filter", "[0-9]");
	//frmScreen.txtCustGtd.setAttribute("filter", "[0-9]");
	//frmScreen.txtGtd.setAttribute("filter", "[0-9]");
	//frmScreen.txtCustReg.setAttribute("filter", "[0-9]");
	//frmScreen.txtReg.setAttribute("filter", "[0-9]");
	frmScreen.txtCustSharesHeld.setAttribute("filter", "[0-9]");
	frmScreen.txtSharesHeld.setAttribute("filter", "[0-9]");
	
	//START: MAR36 - New code added by Joyce Joseph on 11-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.txtCustJobTitle.style.textTransform = "capitalize";
	frmScreen.txtJobTitle.style.textTransform = "capitalize";
	//END: MAR36
	
	<% /* MAR30 Masks for percentage field */ %>
	frmScreen.txtOtherIncomePercentage.setAttribute("filter", "[0-9]");
	frmScreen.txtOtherIncomePercentage.setAttribute("integer", "true");
	frmScreen.txtOtherIncomePercentage.setAttribute("min", "0");
	frmScreen.txtOtherIncomePercentage.setAttribute("max", "100");
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
	//START: MAR119 - New code added by Joyce Joseph on 10-Oct-2005
	//Fields disabled
	frmScreen.optSignedYes.checked=true;
	frmScreen.optNonGuarYes.disabled=true;
	frmScreen.optNonGuarNo.disabled=true;
	frmScreen.optSignedYes.disabled=true;
	frmScreen.optSignedNo.disabled=true;
	//END: MAR119
}
</SCRIPT>

