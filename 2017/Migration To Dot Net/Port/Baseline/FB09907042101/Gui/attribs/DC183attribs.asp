<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtNumberOfYearsEmployed.setAttribute("filter", "[0-9]");
	frmScreen.txtNumberOfMonthsEmployed.setAttribute("filter", "[0-9]");
	frmScreen.txtNumberOfRenewals.setAttribute("filter", "[0-9]"); <%/* MAH 20/11/2006 E2CR35*/%>
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
