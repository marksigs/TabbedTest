<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.cboApplicant.setAttribute("required", "true");
	frmScreen.txtCompanyName.removeAttribute("required");
	frmScreen.txtMonthlyPremium.setAttribute("filter", "[0-9.]");
	frmScreen.txtMonthlyPremium.setAttribute("amount", ".");
	frmScreen.txtMonthlyPremium.setAttribute("min", "0");
	frmScreen.txtMonthlyPremium.setAttribute("max", "999999.99");
	frmScreen.txtYearOfMaturity.setAttribute("integer","true");
	frmScreen.txtYearOfMaturity.setAttribute("filter", "[0-9]");
	var toDay = new Date(); 
	var thisYear = toDay.getFullYear();
	frmScreen.txtYearOfMaturity.setAttribute("min", thisYear.toString());
	frmScreen.txtYearOfMaturity.setAttribute("max", "2100");
	frmScreen.txtMinimumDeathBenefit.setAttribute("filter","[0-9]");
	frmScreen.txtMaturityValue.setAttribute("filter", "[0-9]");
	frmScreen.txtAdditionalDetails.setAttribute("required", "true");
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