<SCRIPT LANGUAGE="JScript">

function SetMasks() 
{
	frmScreen.txtLoanDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtRedemptionDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtRepayment.setAttribute("filter", "[0-9]");
	frmScreen.txtHousePurchase.setAttribute("filter", "[0-9]");
	frmScreen.txtHomeImpr.setAttribute("filter", "[0-9]");
	frmScreen.txtCapitalRaising.setAttribute("filter", "[0-9]");
	frmScreen.txtOrigHousePurchase.setAttribute("filter", "[0-9]");
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