<SCRIPT LANGUAGE="JScript">

function SetMasks()
{	
	frmScreen.txtPurchasePrice.setAttribute("filter", "[0-9]");
	
	frmScreen.txtAmountRequested.setAttribute("filter", "[0-9]");
	frmScreen.txtAmountRequested.setAttribute("required", "y");
	
	frmScreen.txtTermYears.setAttribute("required", "y");
	frmScreen.txtTermYears.setAttribute("filter", "[0-9]");
	
	frmScreen.txtTermMonths.setAttribute("integer", "true");
	frmScreen.txtTermMonths.setAttribute("filter", "[0-9]");
	frmScreen.txtTermMonths.setAttribute("min", "0");
	frmScreen.txtTermMonths.setAttribute("max", "11");
	
	frmScreen.txtInterestOnlyCost.setAttribute("mask", "99999.99");
	frmScreen.txtCapitalAndInterestCost.setAttribute("mask", "99999.99");
	frmScreen.txtPartAndPartCost.setAttribute("mask", "99999.99");

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