<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.cboApplicant.setAttribute("required", "true");
	frmScreen.cboCardType.setAttribute("required", "true");
	frmScreen.txtTotalOutstandingBalance.setAttribute("filter", "[0-9]");
	frmScreen.txtAverageMonthlyRepayment.setAttribute("filter", "[0-9.]");
	frmScreen.txtAverageMonthlyRepayment.setAttribute("amount", ".");
	frmScreen.txtAverageMonthlyRepayment.setAttribute("min", "0");
	frmScreen.txtAverageMonthlyRepayment.setAttribute("max", "999999.99");
	frmScreen.txtAdditionalDetails.setAttribute("required", "true");
	
	//START: MAR36 - New code added by Joyce Joseph on 10-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.txtCardProvider.style.textTransform = "capitalize";
	//END: MAR36
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