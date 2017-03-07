<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtYear1.setAttribute("filter", "[0-9]");
	frmScreen.txtYear2.setAttribute("filter", "[0-9]");
	frmScreen.txtYear3.setAttribute("filter", "[0-9]");
	frmScreen.txtYear1.setAttribute("min", "1900");
	frmScreen.txtYear2.setAttribute("min", "1900");
	frmScreen.txtYear3.setAttribute("min", "1900");
	frmScreen.txtYear1Amount.setAttribute("filter", "[0-9]");
	frmScreen.txtYear2Amount.setAttribute("filter", "[0-9]");
	frmScreen.txtYear3Amount.setAttribute("filter", "[0-9]");
	frmScreen.txtNetMonthlyIncome.setAttribute("filter", "[0-9]");
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
