<SCRIPT LANGUAGE="JScript">

function SetMasks() 
{
	frmScreen.txtCustDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtCustMonthlyRental.setAttribute("filter", "[0-9]");
	frmScreen.txtMonthlyRental.setAttribute("filter", "[0-9]");
	frmScreen.txtArrears.setAttribute("filter", "[0-9]");
	frmScreen.txtMonths.setAttribute("filter", "[0-9]");
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