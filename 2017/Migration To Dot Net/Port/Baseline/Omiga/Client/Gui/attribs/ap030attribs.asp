<SCRIPT LANGUAGE="JScript">

function SetMasks() 
{
	frmScreen.txtTermYears.setAttribute("filter", "[0-9]");
	frmScreen.txtTermMonths.setAttribute("filter", "[0-9]");
	frmScreen.txtTotalLoanAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtOutstandingLoans.setAttribute("filter", "[0-9]");
	frmScreen.txtTotal.setAttribute("filter", "[0-9]");
	frmScreen.txtPurchasePrice.setAttribute("filter", "[0-9]");
	frmScreen.txtPresentValue.setAttribute("filter", "[0-9]");
	frmScreen.txtPostValue.setAttribute("filter", "[0-9]");
	frmScreen.txtCalculator.setAttribute("filter", "[0-9]");
	frmScreen.txtLTV.setAttribute("filter", "[0-9]");
	frmScreen.txtIndemnityAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtIndemnityPremium.setAttribute("filter", "[0-9]");
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